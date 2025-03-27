'환산표 250326-00
이 VBA 매크로는 CSV 파일과 XLSX 파일 간에 등록번호 매칭을 수행하는 도구입니다.
CSV 파일 선택, XLSX 파일 선택, XLSX 파일명 분석하여 처리 방식 결정
데이터 처리 및 결과 CSV 생성하도록 합니다.

코드내용은 다음과 같습니다.
글로벌 변수와 상수 - 모듈 전체에서 필요한 변수와 상수 정의
초기화와 기본 기능 - 워크시트 초기화, 파일 선택, 키워드 감지 기능
유틸리티 함수 - 값 변환, 등록번호 처리 등 공통 함수
파일 처리 함수 - CSV 읽기, 결과 파일 생성 등 파일 관련 유틸리티
처리 모드별 함수 - 각 모드별 처리 함수 (ProcessFilesForStandard, ProcessFilesForSyukei, ProcessFilesForBunseki, ProcessFilesForSyori)
메인 실행 함수 - OneClickProcess로 전체 흐름 제어

그런데 ProcessFilesForBunseki 함수는 XLSX 파일 두 개를 읽고 처리해야하는데, 
금의 코드에서는 하나만 처리하고 있는 것으로 보인다.
읽어들인 두 파일을 처리하여 하나의 파일로 결과를 작성하도록 한다.
이 때 결과의 순서는 fValue의 알파벳순서로 하고 싶다.
