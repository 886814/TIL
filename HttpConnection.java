package honghong;

import java.io.File;
import java.io.FileInputStream;
import java.net.URL;
import java.net.URI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.springframework.http.RequestEntity;
import org.springframework.http.ResponseEntity;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.util.UriComponentsBuilder;


public class HttpConnection {

    public static String filePath = "C:/workspace/";
    public static String fileNm = "스타뱅킹_질문리스트.xlsx";

    public static void main(String[] args) {
        qaPost();

        public static void qaPost () thors Exception {

            try {
                // 엑셀에서 데이터 읽어오기
                FileInputStream file = new FileInputSTream(new File(filePath, fileName));
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                XSSFSheet sheet = workbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();
                int rowIndex = 1;

                // row 한줄씩 읽기
                while (rowIterator.hasNext()){
                    Row row = rowIterator.next();
                    Cell cell = row.getCell(4); // 5번째 컬럼 읽기

                    //서버에 데이터 보내기
                    String url = "http://10.38.136.119.8060/";
                    URI uri = UriComponentsBuilder.fromHttpUrl(url).path("api/func/product/kbProductQa").queryParam("pageSize", "5").queryParam("searchWord", cell.getStringCellValue()).queryParam("svcDstcd", "02")
                            .encode().build().toUri();

                    //서버에서 데이터 받기
                    RestTemplate restTemplate = new RestTemplate();
                    RequestEntity<Void> req = RequestEntity.get(uri).build();
                    ResponseEntity<String> result = restTemplate.exchange(req,String.class);

                    JSONParser parser = new JSONParser();
                    JSONObject jsonObj = (JSONObject) parser.parse(result.getBody());
                    JSONObject json = (JSONObject) jsonObj.get("data");

                    try{
                        JSONArray jsonArray = (JSONArray) json.get("resultList");
                        String res = "";
                        for (int i=0; i<jsonArray.size();i++){
                            JSONObject objectInArray = (JSONObject) jsonArray.get(i);
                            res += ", " + objectInArray.get("brandPrdtCd") + "|" + objectInArray.get("prdctName");
                        }
                        res = res.substring(2);
                        sheet.getRow(rowIndex).createCell(5).setCellValue(res);
                    } catch (Exception e){
                        sheet.getRow(rowIndex).createCell(5).setCellValue("답변 없음");
                    }
                } file.close();
            } catch (Exception e){
                throw e;
        }
    }
}