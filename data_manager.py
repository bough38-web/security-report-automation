# data_manager.py
import os
import json
import zipfile
import pandas as pd
import logging
from datetime import datetime
from config import BASE_DIR, OUTPUT_DIR, KEYWORD_FILE, INITIAL_KEYWORDS

# 로깅 설정
logger = logging.getLogger(__name__)

def setup_environment():
    """출력 디렉토리 생성 및 키워드 파일 초기화"""
    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        if not os.path.exists(KEYWORD_FILE):
            with open(KEYWORD_FILE, "w", encoding="utf-8") as f:
                json.dump(INITIAL_KEYWORDS, f, ensure_ascii=False, indent=2)
            logger.info("초기 keywords.json 파일 생성 완료.")
    except Exception as e:
        logger.error(f"환경 설정 오류: {e}")
        raise

def load_keywords():
    """키워드 파일 로드"""
    try:
        with open(KEYWORD_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        logger.warning(f"{KEYWORD_FILE} 파일이 없어 초기 키워드를 로드합니다.")
        return INITIAL_KEYWORDS
    except json.JSONDecodeError as e:
        logger.error(f"키워드 파일 JSON 디코딩 오류: {e}")
        return INITIAL_KEYWORDS
    except Exception as e:
        logger.error(f"키워드 로드 중 오류 발생: {e}")
        raise

def save_keywords(data):
    """키워드 파일 저장"""
    try:
        with open(KEYWORD_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
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