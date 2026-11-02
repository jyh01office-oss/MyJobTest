# 스탠다드 뱅크 인증 프로젝트 대시보드

Streamlit 기반 교육용 인증 프로젝트 접수 및 진행현황 대시보드입니다. 포함된 CSV는 가상 데모 데이터이며 실제 업무 시스템이나 공식 보고서가 아닙니다.

## 실행

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

`data/SB_project_demo.csv`를 앱과 함께 제공하며, 화면에서 다른 CSV 파일을 업로드할 수도 있습니다. 시험분야 카테고리와 데이터 변경은 현재 Streamlit 세션에만 보관됩니다.
