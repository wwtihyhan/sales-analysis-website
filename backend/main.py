"""
销售数据分析平台 - FastAPI 后端服务
=====================================

主应用入口文件，提供以下API：
- GET  /                    前端页面
- POST /api/upload          上传并分析 Excel 文件（使用 MiniMax AI）
- GET  /api/records         获取历史分析记录
- GET  /api/report/{id}     在线查看报告
- GET  /api/report/{id}/download 下载报告

环境变量配置：
- MINIMAX_API_KEY: MiniMax API 密钥（必填）
- PORT: 服务端口（默认8000）
- DATABASE_URL: PostgreSQL 连接串（可选，默认使用SQLite）

作者: hanyang
日期: 2026-06-18 (创建) ~ 2026-06-21 (AI集成)
"""

import os
import uuid
import logging
from datetime import datetime
from typing import Optional

from fastapi import FastAPI, UploadFile, File, HTTPException, Depends, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, Response
from sqlalchemy.orm import Session
from urllib.parse import quote

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger(__name__)

# 导入数据库模块
from database import SessionLocal, engine, Base, AnalysisRecord

# 创建数据库表
Base.metadata.create_all(bind=engine)

# 创建 FastAPI 应用实例
app = FastAPI(
    title="销售数据分析平台 API",
    description="基于 MiniMax 大模型的智能销售数据分析系统",
    version="2.0.0-AI",  # 版本号标记为AI版本
)

# CORS 中间件配置（允许跨域请求）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 生产环境应限制具体域名
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 挂载静态资源目录
app.mount("/static", StaticFiles(directory="static"), name="static")


# ===== 全局配置 =====

# MiniMax API Key（从环境变量读取）
MINIMAX_API_KEY = os.getenv("MINIMAX_API_KEY")
if not MINIMAX_API_KEY:
    logger.warning("⚠️ 未设置环境变量 MINIMAX_API_KEY，AI分析功能将不可用！")

# 默认金属价格
DEFAULT_GOLD_PRICE = float(os.getenv("DEFAULT_GOLD_PRICE", "930"))
DEFAULT_SILVER_PRICE = float(os.getenv("DEFAULT_SILVER_PRICE", "17"))

logger.info(f"服务初始化完成 | 金价: ¥{DEFAULT_GOLD_PRICE}/g | 银价: ¥{DEFAULT_SILVER_PRICE}/g")


# ===== 数据库会话管理 =====


def get_db():
    """获取数据库会话（依赖注入用）"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


# ===== API 路由 =====


@app.get("/")
async def root():
    """服务首页 - 返回前端HTML页面"""
    return FileResponse("static/index.html")


@app.post("/api/upload")
async def upload_file(
    file: UploadFile = File(...),
    gold_price: Optional[float] = Form(DEFAULT_GOLD_PRICE),
    silver_price: Optional[float] = Form(DEFAULT_SILVER_PRICE),
    db: Session = Depends(get_db),
):
    """
    上传并分析 Excel 文件

    Args:
        file: Excel 文件（.xls/.xlsx）
        gold_price: 当前金价（元/克），默认930
        silver_price: 当前银价（元/克），默认17
        db: 数据库会话（自动注入）

    Returns:
        JSON 格式的分析结果摘要

    Raises:
        HTTPException: 文件格式不支持或分析失败时抛出
    """
    logger.info(f"📤 收到上传请求 | 文件: {file.filename} | 金价: ¥{gold_price} | 银价: ¥{silver_price}")

    # 验证文件格式
    if not file.filename or not file.filename.endswith((".xls", ".xlsx")):
        raise HTTPException(status_code=400, detail="只支持 .xls 或 .xlsx 文件")

    # 验证 API Key
    if not MINIMAX_API_KEY:
        raise HTTPException(
            status_code=503,
            detail="AI分析服务未配置：请设置环境变量 MINIMAX_API_KEY",
        )

    # 保存上传的文件
    file_id = str(uuid.uuid4())
    uploads_dir = "uploads"
    os.makedirs(uploads_dir, exist_ok=True)
    file_path = f"{uploads_dir}/{file_id}_{file.filename}"

    try:
        content = await file.read()
        with open(file_path, "wb") as f:
            f.write(content)
        logger.info(f"✅ 文件已保存 | 路径: {file_path} | 大小: {len(content)} 字节")

    except Exception as e:
        logger.error(f"❌ 文件保存失败: {e}")
        raise HTTPException(status_code=500, detail=f"文件保存失败: {str(e)}")

    # 调用 AI 分析
    try:
        from ai_analyzer import analyze_sales_data_with_ai

        logger.info("🤖 开始 AI 分析...")
        start_time = datetime.now()

        result = analyze_sales_data_with_ai(
            file_path=file_path,
            api_key=MINIMAX_API_KEY,
            gold_price=gold_price or DEFAULT_GOLD_PRICE,
            silver_price=silver_price or DEFAULT_SILVER_PRICE,
        )

        elapsed = (datetime.now() - start_time).total_seconds()
        logger.info(f"✅ AI 分析完成 | 耗时: {elapsed:.1f}s")

    except ImportError as e:
        logger.error(f"❌ AI分析模块加载失败: {e}")
        raise HTTPException(status_code=500, detail="AI分析模块未正确安装")
    except Exception as e:
        logger.error(f"❌ AI 分析失败: {e}")
        raise HTTPException(status_code=500, detail=f"AI分析失败: {str(e)}")

    # 保存到数据库
    try:
        record = AnalysisRecord(
            id=file_id,
            file_name=file.filename,
            upload_time=datetime.now(),
            total_revenue=result.get("total_revenue", 0),
            total_profit=result.get("total_profit", 0),
            report_html=result.get("report_html", ""),
        )
        db.add(record)
        db.commit()
        logger.info(f"💾 记录已保存到数据库 | ID: {file_id}")

    except Exception as e:
        logger.error(f"❌ 数据库写入失败: {e}")
        # 不抛出异常，让用户仍能看到结果

    # 返回结果
    response_data = {
        "id": file_id,
        "file_name": file.filename,
        "total_revenue": result.get("total_revenue", 0),
        "total_profit": result.get("total_profit", 0),
        "report_url": f"/api/report/{file_id}",
        "analysis_method": "ai",  # 标记使用AI分析
    }

    logger.info(f"📤 返回分析结果 | ID: {file_id}")
    return response_data


@app.get("/api/records")
async def get_records(page: int = 1, limit: int = 10, db: Session = Depends(get_db)):
    """
    获取历史分析记录列表（分页）

    Args:
        page: 页码（从1开始）
        limit: 每页记录数（最大50）
        db: 数据库会话

    Returns:
        包含 records 列表和总数的JSON
    """
    from sqlalchemy import desc

    limit = min(limit, 50)  # 限制每页最多50条
    offset = (page - 1) * limit

    records = (
        db.query(AnalysisRecord)
        .order_by(desc(AnalysisRecord.upload_time))
        .offset(offset)
        .limit(limit)
        .all()
    )
    total = db.query(AnalysisRecord).count()

    return {
        "records": [
            {
                "id": r.id,
                "file_name": r.file_name,
                "upload_time": r.upload_time.strftime("%Y-%m-%d %H:%M"),
                "total_revenue": r.total_revenue,
                "total_profit": r.total_profit,
            }
            for r in records
        ],
        "total": total,
        "page": page,
        "limit": limit,
    }


@app.get("/api/report/{record_id}")
async def get_report(record_id: str, db: Session = Depends(get_db)):
    """
    在线查看分析报告（HTML格式）

    Args:
        record_id: 分析记录ID
        db: 数据库会话

    Returns:
        HTMLResponse（直接在浏览器渲染）
    """
    record = db.query(AnalysisRecord).filter(AnalysisRecord.id == record_id).first()

    if not record:
        raise HTTPException(status_code=404, detail="记录不存在")

    if not record.report_html:
        raise HTTPException(status_code=404, detail="报告内容为空")

    return HTMLResponse(content=record.report_html, media_type="text/html")


@app.get("/api/report/{record_id}/download")
async def download_report(record_id: str, db: Session = Depends(get_db)):
    """
    下载分析报告（HTML文件）

    Args:
        record_id: 分析记录ID
        db: 数据库会话

    Returns:
        Response（触发浏览器文件下载）
    """
    record = db.query(AnalysisRecord).filter(AnalysisRecord.id == record_id).first()

    if not record:
        raise HTTPException(status_code=404, detail="记录不存在")

    safe_filename = quote(f"sales_report_{record.file_name}_{record.id[:8]}.html")
    headers = {
        "Content-Disposition": f'attachment; filename="{safe_filename}"',
    }

    return Response(
        content=record.report_html,
        media_type="text/html; charset=utf-8",
        headers=headers,
    )


@app.get("/api/health")
async def health_check():
    """健康检查接口（用于监控）"""
    status = {
        "status": "ok",
        "version": "2.0.0-AI",
        "timestamp": datetime.now().isoformat(),
        "ai_enabled": bool(MINIMAX_API_KEY),
        "default_gold_price": DEFAULT_GOLD_PRICE,
        "default_silver_price": DEFAULT_SILVER_PRICE,
    }
    return status


# ===== 应用启动事件 =====


@app.on_event("startup")
async def startup_event():
    """应用启动时执行"""
    logger.info("=" * 60)
    logger.info("🚀 销售数据分析平台 启动中...")
    logger.info(f"   版本: 2.0.0-AI (MiniMax大模型驱动)")
    logger.info(f"   AI功能: {'✅ 已启用' if MINIMAX_API_KEY else '❌ 未启用'}")
    logger.info("=" * 60)


if __name__ == "__main__":
    import uvicorn

    port = int(os.getenv("PORT", 8000))
    logger.info(f"启动本地开发服务器... 端口: {port}")
    uvicorn.run(app, host="0.0.0.0", port=port, reload=True)
