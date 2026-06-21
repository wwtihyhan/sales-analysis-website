"""
MiniMax API Client
==================

封装 MiniMax 大模型 API 调用逻辑，支持同步和异步调用。

官方文档: https://platform.minimaxi.com/document/Chatcompletion_v2

支持的模型:
- MiniMax-M3 (最新，推荐)
- MiniMax-M2.7 / M2.7-highspeed
- MiniMax-M2.5 / M2.5-highspeed
- MiniMax-M2.1
- MiniMax-M2

作者: hanyang
日期: 2026-06-21
"""

import os
import json
import time
import logging
from typing import Optional, Dict, Any, List

import httpx

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class MiniMaxClient:
    """MiniMax 大模型 API 客户端"""

    # API 基础配置
    BASE_URL = "https://api.minimaxi.com/v1"
    CHAT_ENDPOINT = "/text/chatcompletion_v2"

    # 默认参数
    DEFAULT_MODEL = "MiniMax-M3"
    DEFAULT_TEMPERATURE = 0.7
    DEFAULT_MAX_TOKENS = 16384
    DEFAULT_TIMEOUT = 120.0  # 秒（分析任务可能较慢）

    # 重试配置
    MAX_RETRIES = 3
    RETRY_DELAY = 2.0  # 秒

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: str = DEFAULT_MODEL,
        temperature: float = DEFAULT_TEMPERATURE,
        max_tokens: int = DEFAULT_MAX_TOKENS,
        timeout: float = DEFAULT_TIMEOUT,
    ):
        """
        初始化 MiniMax 客户端

        Args:
            api_key: API 密钥（如果为空则从环境变量 MINIMAX_API_KEY 读取）
            model: 模型名称（默认 MiniMax-M3）
            temperature: 温度系数 (0, 1]
            max_tokens: 最大生成 token 数
            timeout: 请求超时时间（秒）
        """
        self.api_key = api_key or os.getenv("MINIMAX_API_KEY")
        if not self.api_key:
            raise ValueError(
                "API Key 未提供！请通过参数传入或设置环境变量 MINIMAX_API_KEY"
            )

        self.model = model
        self.temperature = temperature
        self.max_tokens = max_tokens
        self.timeout = timeout

        # 创建 HTTP 客户端
        self.client = httpx.Client(timeout=timeout)

        logger.info(f"MiniMax 客户端初始化完成 | 模型: {self.model}")

    def _build_headers(self) -> Dict[str, str]:
        """构建请求头"""
        return {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }

    def _build_request_body(
        self,
        messages: List[Dict[str, str]],
        stream: bool = False,
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
    ) -> Dict[str, Any]:
        """
        构建请求体

        Args:
            messages: 对话消息列表 [{"role": "user", "content": "..."}]
            stream: 是否使用流式输出
            temperature: 温度系数（可选，覆盖默认值）
            max_tokens: 最大 token 数（可选，覆盖默认值）

        Returns:
            请求体字典
        """
        body = {
            "model": self.model,
            "messages": messages,
            "stream": stream,
            "max_completion_tokens": max_tokens or self.max_tokens,
            "temperature": temperature or self.temperature,
        }

        return body

    def chat(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
    ) -> str:
        """
        发送对话请求（同步）

        Args:
            prompt: 用户输入的提示词
            system_prompt: 系统提示词（可选）
            temperature: 温度系数（可选）
            max_tokens: 最大 token 数（可选）

        Returns:
            模型的回复内容（字符串）

        Raises:
            Exception: API 调用失败时抛出异常
        """
        # 构建消息列表
        messages = []

        # 添加系统提示词
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})

        # 添加用户消息
        messages.append({"role": "user", "content": prompt})

        # 构建请求
        url = f"{self.BASE_URL}{self.CHAT_ENDPOINT}"
        headers = self._build_headers()
        body = self._build_request_body(messages, temperature=temperature, max_tokens=max_tokens)

        logger.info(f"发送请求到 MiniMax API | 模型: {self.model} | Prompt 长度: {len(prompt)} 字符")

        # 重试机制
        last_exception = None
        for attempt in range(1, self.MAX_RETRIES + 1):
            try:
                logger.info(f"第 {attempt}/{self.MAX_RETRIES} 次尝试...")

                response = self.client.post(url, headers=headers, json=body)
                response.raise_for_status()

                result = response.json()

                # 解析响应
                content = self._parse_response(result)

                logger.info(
                    f"API 调用成功 | Token 使用: {result.get('usage', {}).get('total_tokens', 'N/A')}"
                )

                return content

            except httpx.HTTPStatusError as e:
                error_msg = self._extract_error_message(e.response)
                last_exception = Exception(f"HTTP {e.response.status_code}: {error_msg}")
                logger.error(f"请求失败 (尝试 {attempt}): {error_msg}")

                # 如果是认证错误或频率限制，不重试
                if e.response.status_code in [401, 403, 429]:
                    raise last_exception

            except httpx.TimeoutException as e:
                last_exception = Exception(f"请求超时 ({self.timeout}秒)")
                logger.error(f"请求超时 (尝试 {attempt}): {e}")

            except json.JSONDecodeError as e:
                last_exception = Exception(f"响应解析失败: {e}")
                logger.error(f"JSON 解析失败 (尝试 {attempt}): {e}")

            except Exception as e:
                last_exception = e
                logger.error(f"未知错误 (尝试 {attempt}): {e}")

            # 等待后重试
            if attempt < self.MAX_RETRIES:
                wait_time = self.RETRY_DELAY * attempt
                logger.info(f"等待 {wait_time} 秒后重试...")
                time.sleep(wait_time)

        # 所有重试都失败
        raise last_exception or Exception("API 调用失败：已达到最大重试次数")

    def chat_with_json(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        temperature: Optional[float] = None,
    ) -> Dict[str, Any]:
        """
        发送对话请求并返回 JSON 结果

        与 chat() 类似，但会自动将回复内容解析为 JSON 格式。
        适用于需要结构化输出的场景（如数据分析）。

        Args:
            prompt: 用户输入的提示词（应包含 JSON 输出要求）
            system_prompt: 系统提示词（可选）
            temperature: 温度系数（可选，建议降低以获得更确定的结果）

        Returns:
            解析后的 JSON 字典

        Raises:
            json.JSONDecodeError: 回复内容无法解析为有效 JSON
            Exception: API 调用失败
        """
        # 使用较低的温度以确保输出格式稳定
        actual_temp = temperature or min(self.temperature, 0.5)

        # 调用 API
        raw_response = self.chat(
            prompt=prompt,
            system_prompt=system_prompt,
            temperature=actual_temp,
        )

        # 尝试从响应中提取 JSON
        parsed_json = self._extract_json_from_response(raw_response)

        logger.info(f"成功解析 JSON 响应 | 键数量: {len(parsed_json)}")

        return parsed_json

    def _parse_response(self, response_data: Dict[str, Any]) -> str:
        """
        解析 API 响应，提取文本内容

        Args:
            response_data: API 返回的完整 JSON 数据

        Returns:
            提取到的文本内容

        Raises:
            KeyError: 响应格式不符合预期
        """
        # 检查错误码
        base_resp = response_data.get("base_resp", {})
        status_code = base_resp.get("status_code", 0)
        if status_code != 0:
            status_msg = base_resp.get("status_msg", "未知错误")
            raise Exception(f"API 返回错误 [{status_code}]: {status_msg}")

        # 提取 choices
        choices = response_data.get("choices")
        if not choices or len(choices) == 0:
            raise Exception("API 响应为空：没有 choices")

        # 提取 message.content
        message = choices[0].get("message", {})
        content = message.get("content", "")

        if not content:
            raise Exception("API 响应为空：message.content 为空")

        return content

    @staticmethod
    def _extract_error_message(response: httpx.Response) -> str:
        """
        从错误响应中提取可读的错误信息

        Args:
            response: HTTP 错误响应对象

        Returns:
            错误信息字符串
        """
        try:
            data = response.json()
            # 尝试提取 MiniMax 的错误信息
            base_resp = data.get("base_resp", {})
            if base_resp.get("status_msg"):
                return base_resp["status_msg"]
            # 尝试通用错误信息
            if data.get("detail"):
                return data["detail"]
            if data.get("error", {}).get("message"):
                return data["error"]["message"]
            return json.dumps(data, ensure_ascii=False)[:200]
        except Exception:
            return response.text[:200]

    @staticmethod
    def _extract_json_from_response(text: str) -> Dict[str, Any]:
        """
        从响应文本中提取并解析 JSON

        支持处理以下情况：
        - 纯 JSON 文本
        - 包含在 markdown 代码块中的 JSON
        - JSON 前后有其他文字说明

        Args:
            text: 原始响应文本

        Returns:
            解析后的 JSON 字典
        """
        text = text.strip()

        # 尝试直接解析
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass

        # 尝试查找 ```json ... ``` 代码块
        import re
        json_pattern = r'```(?:json)?\s*([\s\S]*?)```'
        matches = re.findall(json_pattern, text)

        for match in matches:
            match = match.strip()
            try:
                return json.loads(match)
            except json.JSONDecodeError:
                continue

        # 尝试查找第一个 { ... } 或 [ ... ]
        brace_pattern = r'\{[\s\S]*\}'
        brace_matches = re.findall(brace_pattern, text)

        for match in brace_matches:
            try:
                return json.loads(match)
            except json.JSONDecodeError:
                continue

        # 所有方法都失败，抛出异常
        raise json.JSONDecodeError(
            f"无法从响应中提取有效的 JSON。\n原始响应前500字符:\n{text[:500]}",
            text,
            0,
        )

    async def chat_async(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        **kwargs,
    ) -> str:
        """
        发送对话请求（异步版本）

        与 chat() 功能相同，但支持异步调用。
        适用于 FastAPI 等异步框架。

        Args:
            prompt: 用户输入的提示词
            system_prompt: 系统提示词（可选）
            **kwargs: 其他参数（传递给 chat 方法）

        Returns:
            模型的回复内容（字符串）
        """
        # 在实际应用中可以使用 httpx.AsyncClient
        # 这里为了简化，先使用同步调用
        return self.chat(prompt=prompt, system_prompt=system_prompt, **kwargs)

    def close(self):
        """关闭 HTTP 客户端"""
        if self.client:
            self.client.close()
            logger.info("MiniMax 客户端已关闭")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


# ===== 工厂函数 & 便捷方法 =====


def create_client(
    api_key: Optional[str] = None,
    model: str = "MiniMax-M3",
) -> MiniMaxClient:
    """
    创建 MiniMax 客户端的便捷工厂函数

    Args:
        api_key: API 密钥（可选，默认从环境变量读取）
        model: 模型名称

    Returns:
        配置好的 MiniMaxClient 实例

    Example:
        >>> client = create_client(model="MiniMax-M2.5")
        >>> result = client.chat("你好")
        >>> print(result)
    """
    return MiniMaxClient(api_key=api_key, model=model)


def test_connection(api_key: Optional[str] = None) -> bool:
    """
    测试 API 连接是否正常

    Args:
        api_key: API 密钥（可选）

    Returns:
        True 表示连接正常，False 表示连接失败
    """
    try:
        with create_client(api_key=api_key) as client:
            response = client.chat("说 'OK'")
            return "OK" in response or len(response) > 0
    except Exception as e:
        logger.error(f"连接测试失败: {e}")
        return False


if __name__ == "__main__":
    # 简单测试
    print("=" * 60)
    print("MiniMax API 客户端测试")
    print("=" * 60)

    # 从环境变量读取 API Key
    api_key = os.getenv("MINIMAX_API_KEY")
    if not api_key:
        print("❌ 请设置环境变量 MINIMAX_API_KEY")
        exit(1)

    print(f"\n✅ API Key 已加载: {api_key[:10]}...{api_key[-6:]}")

    # 创建客户端
    with create_client(api_key=api_key) as client:
        # 测试简单对话
        print("\n📤 发送测试请求...")
        start_time = time.time()
        response = client.chat("用一句话介绍你自己")
        elapsed = time.time() - start_time

        print(f"\n📥 收到响应 (耗时 {elapsed:.2f}s):")
        print("-" * 40)
        print(response)
        print("-" * 40)

        # 显示 Token 使用情况
        print("\n✅ 测试完成！")

    print("\n💡 提示：在生产环境中，请将 API Key 设置为环境变量 MINIMAX_API_KEY")
