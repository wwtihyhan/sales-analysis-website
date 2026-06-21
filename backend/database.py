from sqlalchemy import create_engine, Column, String, Float, DateTime, Text
from sqlalchemy.orm import declarative_base, sessionmaker
from datetime import datetime
import os

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./sales_analysis.db")

# 如果是 postgres:// 开头，替换为 pg8000 驱动
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+pg8000://", 1)
if DATABASE_URL.startswith("postgresql://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+pg8000://", 1)

engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


class AnalysisRecord(Base):
    __tablename__ = "analysis_records"
    id = Column(String, primary_key=True)
    file_name = Column(String, nullable=False)
    upload_time = Column(DateTime, default=datetime.now)
    total_revenue = Column(Float)
    total_profit = Column(Float)
    report_html = Column(Text)
    created_at = Column(DateTime, default=datetime.now)


def create_tables():
    Base.metadata.create_all(bind=engine)
