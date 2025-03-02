file읽기 코드 비교

Dim stream As Object 
Set stream = CreateObject("ADODB.Stream") 
On Error Resume Next 
stream.Charset = "UTF-8" ' まずUTF-8で試す 
stream.Open 
stream.LoadFromFile filePath
ReadTextFileContent = stream.ReadText 
stream.Close 
Set stream = Nothing


fileNum = FreeFile
Open filePath For Binary As #fileNum
ReDim byteData(LOF(fileNum) - 1)
Get #fileNum, , byteData
Close #fileNum

allText = StrConv(byteData, vbUnicode)
fileContent = allText

lines = Split(fileContent, vbCrLf)

첫 번째 코드: 기본 VBA 파일 I/O 함수를 사용하여 바이너리 모드로 파일을 읽습니다.
두 번째 코드: ADODB.Stream 객체를 사용하여 텍스트 파일을 읽습니다.

네, VBA에서 파일을 읽는 두 가지 다른 방법에 대해 설명해 드리겠습니다.

ADODB.Stream을 사용한 코드 설명

Dim stream As Object '파일 데이터를 처리할 ADODB.Stream 객체 변수를 선언합니다.
Set stream = CreateObject("ADODB.Stream") 'ADODB.Stream 객체를 생성하여 stream 변수에 할당합니다.
On Error Resume Next '오류가 발생해도 코드 실행을 계속하도록 설정합니다.

stream.Charset = "UTF-8"  ' 스트림의 문자셋을 UTF-8로 설정합니다. 
stream.Open '스트림을 엽니다.
stream.LoadFromFile filePath 'filePath에 지정된 파일을 스트림에 로드합니다.

ReadTextFileContent = stream.ReadText '스트림에서 텍스트를 읽어 ReadTextFileContent 변수에 저장합니다.

stream.Close '스트림을 닫습니다.

Set stream = Nothing '스트림 객체 참조를 해제합니다.



기본 VBA 파일 I/O 함수를 사용한 코드 설명

fileNum = FreeFile '사용 가능한 파일 번호를 가져와 fileNum 변수에 저장합니다.
Open filePath For Binary As #fileNum 'filePath 경로의 파일을 바이너리 모드로 엽니다.
ReDim byteData(LOF(fileNum) - 1) 'LOF(fileNum)은 파일 크기를 반환하고, 그만큼 byteData 배열을 재정의합니다.
Get #fileNum, , byteData '파일에서 데이터를 읽어 byteData 배열에 저장합니다.
Close #fileNum '파일을 닫습니다.

allText = StrConv(byteData, vbUnicode) 'byteData 바이트 배열을 유니코드 문자열로 변환합니다.

fileContent = allText '변환된 텍스트를 fileContent 변수에 할당합니다.

lines = Split(fileContent, vbCrLf) 'fileContent 문자열을 줄바꿈 문자(vbCrLf)로 분할하여 lines 배열에 저장합니다.



두 방법의 차이점

사용하는 기술:

첫 번째 방법: Microsoft ActiveX Data Objects(ADODB) 라이브러리의 Stream 객체 사용
두 번째 방법: VBA의 기본 파일 입출력 함수 사용


인코딩 처리:

ADODB.Stream: 문자셋(Charset)을 명시적으로 지정 가능(예: UTF-8)
기본 VBA I/O: 바이너리로 읽은 후 StrConv 함수로 변환, 특정 인코딩을 직접 지정하기 어려움


오류 처리:

ADODB.Stream: On Error Resume Next로 오류를 처리하고 파일 크기 검사
기본 VBA I/O: 별도의 오류 처리가 코드에 포함되어 있지 않음


유연성:

ADODB.Stream: 다양한 인코딩 지원, 바이너리/텍스트 모드 전환 가능
기본 VBA I/O: 단순하지만 인코딩 처리가 제한적


속도와 메모리:

ADODB.Stream: 대용량 파일 처리에 유리할 수 있음
기본 VBA I/O: 작은 파일에는 충분하지만 큰 파일 처리 시 메모리 문제 발생 가능


코드 복잡성:

ADODB.Stream: 더 많은 코드 라인, 오류 처리 포함
기본 VBA I/O: 상대적으로 간단한 코드 구조


두 방식 모두 파일을 읽는 목적은 같지만, 
인코딩 처리와 오류 관리 측면에서 ADODB.Stream 방식이 더 견고하고 다양한 인코딩을 지원하는 장점이 있습니다.