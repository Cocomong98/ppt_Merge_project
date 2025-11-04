# 📚 PPTX 순서 병합 프로그램

PyQt5 GUI 기반의 파이썬 애플리케이션으로, 여러 개의 파워포인트 파일을 사용자가 지정한 순서대로 하나의 `PPTX` 파일로 병합합니다. Windows 환경에서 실행 파일(`EXE`)로 배포하는 것을 목표로 합니다.

## ✨ 주요 기능

* **GUI 기반:** 직관적인 PyQt5 사용자 인터페이스 제공.
* **파일 관리:** 드래그 앤 드롭 또는 파일 탐색기를 통해 PPT/PPTX 파일 추가 및 순서 변경, 제거 가능.
* **자동 변환 지원 (Windows 전용):** 구형 `.ppt` 파일을 병합이 가능한 `.pptx` 형식으로 자동 변환합니다. (**MS PowerPoint 설치 필수**)
* **비동기 처리:** 병합 작업 중에도 GUI가 멈추지 않도록 스레드를 사용합니다.

## 🛠️ 설치 및 실행

### 1. 개발 환경 설정 (Windows 권장)

이 프로젝트는 Windows 환경에서 최종 패키징 되어야 합니다.

```bash
# 1. 가상 환경 생성 및 활성화
python -m venv venv
.\venv\Scripts\Activate

# 2. 필수 라이브러리 설치
pip install -r requirements.txt