# 수입원가 계산기 배포 방법

이 앱은 `shipping list` 엑셀과 프렌드 해운항공 `INV` PDF를 업로드해서 수입원가를 계산하는 웹앱입니다.

## Render 배포

1. GitHub에 이 폴더를 업로드합니다.
2. Render에서 `New Web Service`를 선택합니다.
3. GitHub 저장소를 연결합니다.
4. 아래 값으로 배포합니다.
   - Environment: `Python`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python server.py`
   - Health Check Path: `/health`
5. 배포가 끝나면 Render가 만들어주는 `https://...onrender.com` 주소를 공유하면 됩니다.

## 로컬 실행

```bash
python server.py
```

브라우저에서 아래 주소를 엽니다.

```text
http://127.0.0.1:8765
```

## 데이터 저장 방식

업로드한 엑셀/PDF 파일은 서버에 저장하지 않고, 읽어서 계산 결과만 브라우저로 돌려줍니다.
임시저장은 각 사용자 브라우저의 `localStorage`에 저장됩니다.

## 공개 사용 전 확인할 점

- 링크를 아는 사람은 누구나 접속할 수 있습니다.
- 거래처 파일을 다루는 서비스라면 로그인 기능을 추가하는 것이 좋습니다.
- 무료 배포 서버는 일정 시간 사용하지 않으면 잠들 수 있습니다.
