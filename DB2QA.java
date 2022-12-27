package honghong;

import java.io.File;
import java.io.FileOutputStream;
import java.net.URI;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.springframework.http.RequestEntity;
import org.springframework.http.ResponseEntity;
import org.springframework.http.HttpStatus;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.util.UriComponentsBuilder;


public class DB2QA {

    public static String filePath = "C:/workspace/";
    public static String fileNm = "스타뱅킹_질문리스트.xlsx";

    public static void main(String[] args) {
        qaPost();
    }

    public static Connection dbConn;

    public static Connection getConnection(){
        Connection conn = null;
        try{
            String user = "아이디";
            String pw = "비번";
            String url = "jdbc:db2://~~";

            Class.forName("com.ibm.db2.jcc.DB2Driver").newInstance();
            conn = DriverManeger.getConnection(url, user, pw);
        } catch (Exception e)
        return conn;
    }
    public static void qaPost () thors Exception {

        try {

            Connection conn = null;
            PreparedStatement pstm = null;
            ResultSet rs = null;

            XSSFSheet sheet = null;
            XSSFRow row = null;
            XSSFCell cell = null;

            // 파일 생성
            XSSFWorkbook workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("sheet1");

            File file = new File(filePath, fileName);
            FileOutputStream fos = new FileOutputStream(file);

            // header 생성
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("질문");
            cell = row.createCell(1);
            cell.setCellValue("답변");

            // DB 조회
            String query = "쿼리 내용";
            conn = DBConnection.getConnection();
            pstm = conn.prepareStatement(query);
            rs = pstm.executeQuery();

            int rowIndex = 1;
            XSSFRow curRow;

            while (rs.next()){
                String question = rs.getString("검색어내용");
                curRow = sheet.createRow(rowIndex);
                curRow.createCell(0).setCellValue(question);

                //서버에 데이터 보내기
                String url = "http://10.38.136.119.8060/";
                URI uri = UriComponentsBuilder.fromHttpUrl(url).path("api/func/product/kbProductQa").queryParam("pageSize", "5").queryParam("searchWord", cell.getStringCellValue()).queryParam("svcDstcd", "02")
                        .encode().build().toUri();

                //서버에서 데이터 받기
                RestTemplate restTemplate = new RestTemplate();
                RequestEntity<Void> req = RequestEntity.get(uri).build();
                ResponseEntity<String> result = restTemplate.exchange(req,String.class);
                HttpStatus responseCode = result.getStatusCode();

                // Data Parsing
                JSONParser parser = new JSONParser();
                JSONObject jsonObj = (JSONObject) parser.parse(result.getBody());
                JSONObject json = (JSONObject) jsonObj.get("data");

                //응답할 경우
                if (responseCode.value() == 200){
                    try{
                        JSONArray jsonArray = (JSONArray) json.get("resultList");
                        String res = "";
                        for (int i=0; i<jsonArray.size();i++){
                            JSONObject objectInArray = (JSONObject) jsonArray.get(i);
                            res += ", " + objectInArray.get("brandPrdtCd") + "|" + objectInArray.get("prdctName");
                        }
                        res = res.substring(2);
                        curRow.createCell(1).setCellValue(res);
                    } catch (Exception e){ // 응답은 하지만 답변이 없는 경우
                        curRow.createCell(1).setCellValue("답변 없음");
                    }
                } // 응답 없을 시 INTERNAL_SERVER_ERROR 등
                else curRow.createCell(1).setCellValue("응답 오류: " + responseCode);
                rowIndex ++;
            }
            rs.close();
            pstm.close();
            conn.close();
            workbook.write(fos);
            fos.close();
            } catch (Exception e){
                    e.printStackTrace();
                }

    }
}