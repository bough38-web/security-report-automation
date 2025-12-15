# report_generator.py
from pptx import Presentation
from config import PPT_TITLE, PPT_FOOTER, PPT_DATE
import logging

logger = logging.getLogger(__name__)

def make_ppt(summary_map: dict, path: str):
    """요약본을 기반으로 PowerPoint 보고서 생성"""
    try:
        prs = Presentation()
        # 1. 표지 슬라이드 (Layout 0: Title Slide)
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = PPT_TITLE
        slide.placeholders[1].text = PPT_DATE

        # 2. 키워드별 요약 슬라이드 (Layout 1: Title and Content)
        for k, v in summary_map.items():
            s = prs.slides.add_slide(prs.slide_layouts[1])
            s.shapes.title.text = f"📰 {k} 뉴스 요약"
            s.placeholders[1].text = v

        # 3. 마무리 슬라이드
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = "✅ 감사 및 보고 대응 안내"
        s.placeholders[1].text = PPT_FOOTER

        prs.save(path)
        logger.info(f"PPT 파일 생성 성공: {path}")
    except Exception as e:
        logger.error(f"PPT 생성 중 오류 발생: {e}")
        raise