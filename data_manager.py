import json
import os
import logging
import zipfile
import pandas as pd
from dotenv import load_dotenv
from config import DATA_FILE

logger = logging.getLogger(__name__)

def setup_environment():
    """환경 변수(.env)를 로드합니다."""
    # GitHub Actions에서는 Secrets가 자동으로 주입되므로 .env가 없어도 괜찮습니다.
    load_dotenv() 

def load_keywords():
    """JSON 파일에서 키워드 목록을 불러옵니다."""
    if not os.path.exists(DATA_FILE):
        return []
    try:
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return []

def save_keywords(keywords):
    """키워드 목록을 JSON 파일에 저장합니다."""
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(keywords, f, ensure_ascii=False, indent=4)
        logger.info("키워드 파일 저장 완료.")
    except Exception as e:
        logger.error(f"키워드 파일 저장 오류: {e}")
        raise

def save_reports(df: pd.DataFrame, summary_map: dict, excel_path: str, ppt_path: str, zip_path: str):
    """Excel, PPT 파일 저장 및 ZIP 압축"""
    from report_generator import make_ppt # 순환 참조 방지를 위해 내부에서 import

    try:
        # 1. Excel 저장
        df.to_excel(excel_path, index=False)
        logger.info(f"Excel 보고서 저장 완료: {excel_path}")

        # 2. PPT 저장
        make_ppt(summary_map, ppt_path)
        logger.info(f"PPT 보고서 저장 완료: {ppt_path}")

        # 3. ZIP 압축
        with zipfile.ZipFile(zip_path, "w") as z:
            z.write(excel_path, os.path.basename(excel_path))
            z.write(ppt_path, os.path.basename(ppt_path))
        logger.info(f"보고서 ZIP 압축 완료: {zip_path}")
        
        return excel_path, ppt_path, zip_path
    except Exception as e:
        logger.error(f"보고서 저장 및 압축 오류: {e}")
        raise
