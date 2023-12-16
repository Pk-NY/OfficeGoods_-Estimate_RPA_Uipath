### [견적서 메일 발송] ###

메일로 수신한 작업지시서를 읽어서 사무용품 사이트에서 물품의 견적서를 작성하여 요청자에게 메일 발송



**[사전작업]**

Config.xlsx : input과 output 폴더 경로, 사이트URL, 메일계정, 메일제목, 메일내용, 결과파일 이름


**[개인작업파일 mywork 폴더]**

   GetAttachmentFile 수신한 메일에 첨부된 작업지시서를 다운
  
   ReadExcel_DataTable : 작업지시서를 읽고 DT생성 (물품코드, 수량)

   CheckOutputFolder : output폴더확인후생성

   GetPdf - 장바구니 페이지에서 견적서를 클릭하여 PDF 파일로 생성후 저장
   
   SendEmail - 저장한 PDF 파일을 확인, 첨부하여 메일 발송
   
   

### [작업과정] ###

1. **INITIALIZE PROCESS**
   
   수신한 메일에 첨부된 작업지시서를 input폴더에 다운받고 읽어서 DT 생성 -> Transactiondata 물품코드, 수량
   
2. **GET TRANSACTION DATA**

    Transaction Item의 타입 Datarow. Transactiondata의 물품코드, 수량
   
3. **PROCESS TRANSACTION**

   사무용품 사이트에서 작업지시서의 물품코드를 검색하여 나온 물품의 수량을 작업지시서 대로 입력하여 장바구니에 담기

    
4. **END PROCESS**
   
   장바구니 페이지에서 견적서를 클릭하여 PDF 파일로 생성후 저장하여 메일에 첨부하여 발송
   
