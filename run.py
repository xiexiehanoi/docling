"""
문서 텍스트 & 이미지 추출기 (Gemini 3.0 Flash Preview)
사용법: python3 run.py 파일경로.pdf
        python3 run.py 파일경로.pptx
        python3 run.py 파일경로.ppt
"""

import sys
import os
import warnings

warnings.filterwarnings("ignore")

from dotenv import load_dotenv
from extractor import GeminiDocumentExtractor

load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY", "")

MODEL = "gemini-3-flash-preview"


def main():
    if len(sys.argv) < 2:
        print("사용법: python3 run.py <파일경로>")
        print("지원 형식: .pdf, .pptx, .ppt")
        sys.exit(1)

    file_path = sys.argv[1]

    if not os.path.exists(file_path):
        print(f"파일을 찾을 수 없습니다: {file_path}")
        sys.exit(1)

    if not API_KEY:
        print(".env 파일에 GEMINI_API_KEY를 설정해주세요.")
        sys.exit(1)

    # 저장 경로 선택
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    default_dir = os.path.join(os.path.dirname(os.path.abspath(file_path)), f"{base_name}_output")

    print(f"\n파일: {file_path}")
    print(f"모델: {MODEL}")
    user_dir = input(f"저장 경로 (Enter = {default_dir}): ").strip()
    output_dir = user_dir if user_dir else default_dir

    print(f"\n출력: {output_dir}/")
    print("-" * 50)

    def on_progress(ratio, msg):
        print(f"  [{int(ratio * 100):3d}%] {msg}")

    try:
        extractor = GeminiDocumentExtractor(API_KEY, MODEL)
        texts, images = extractor.extract(file_path, output_dir, on_progress)
    except Exception as e:
        print(f"\n오류 발생: {e}")
        sys.exit(1)

    print("-" * 50)
    print(f"텍스트: {len(texts)}페이지 추출 → {output_dir}/texts/")
    print(f"이미지: {len(images)}개 추출  → {output_dir}/images/")

    # 추출된 텍스트 화면 출력
    print("\n" + "=" * 50)
    for t in texts:
        print(f"\n── Page {t['page']} {'─' * 40}")
        print(t["text"])
    print("\n" + "=" * 50)
    print(f"\n전체 결과 저장됨: {output_dir}/")


if __name__ == "__main__":
    main()
