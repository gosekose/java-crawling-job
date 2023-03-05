import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

public class Main {

    static int colIdxH = 7;
    static int colIdxI = 8;
    static HashMap<String, Company> companies = new HashMap<>();

    public static void main(String[] args) throws IOException, InterruptedException {

        if (args.length > 0) {
            int startIdx = Integer.parseInt(args[0]);
            writeExcel(startIdx);
        } else {
            System.out.println("시작 인덱스를 입력해주세요.");
        }

//        writeExcel(20);

    }

    private static void writeExcel(int startIdx) throws InterruptedException, IOException {
        for (int i = startIdx; i < 933; i++) {

            if (i % 10 == 0) {
                Thread.sleep(10000);
            }

            FileInputStream file = new FileInputStream("./db2.xlsx");
            Workbook workbook = new XSSFWorkbook(file);

            Sheet sheet = workbook.getSheetAt(1);
            Row row = sheet.getRow(i);

            String companyMenu = row.getCell(4).getStringCellValue();
            String companyName = row.getCell(6).getStringCellValue(); // 셀의 값을 읽음
            System.out.println("순번 = " + i);

            String business = "";
            String classification = "";

            if (companyName.equals("한국전력공사") || companyName.equals("현대건설") ||
                    companyName.equals("서울교통공사") || companyName.equals("LG디스플레이") ||
                    companyName.equals("Sk하이닉스") || companyName.equals("SK하이닉스") ||
                    companyName.contains("현대자동차")) {

                String cls;
                String busi;

                if (companyName.equals("한국전력공사")) {
                    cls = "공공기관";
                    busi = "송전 및 배전업";
                }

                else if (companyName.equals("현대건설")) {
                    cls = "대기업";
                    busi = "종합건설,주택분양,건설산업부문 설계,감리";
                }

                else if (companyName.equals("서울교통공사")) {
                    cls = "공공기관";
                    busi = "도시철도 운송/기관차,철도차량부품 제조/지하철공사";
                }

                else if (companyName.equals("LG디스플레이")) {
                    cls = "대기업";
                    busi = "액정표시장치(TFT-LCD) 제조";
                }

                else if (companyName.equals("Sk하이닉스") || companyName.equals("SK하이닉스")) {
                    cls = "대기업";
                    busi = "반도체,컴퓨터,통신기기 제조,도매";
                }

                else {
                    cls = "대기업";
                    busi = "자동차(승용차,버스,트럭,특장차),자동차부품,자동차전착도료 제조,차량정비사업/항공기,부속품 도소매/별정통신,부가통신/부동산 임대";
                }

                Row writeRow = sheet.getRow(i);
                Cell writeICell = writeRow.getCell(colIdxI, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                writeICell.setCellValue(cls);
                System.out.println("기업 구분 = " + cls);

                Cell writeHCell = writeRow.getCell(colIdxH, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                writeHCell.setCellValue(busi);
                System.out.println("주요 사업 = " + busi);

                FileOutputStream outputStream = new FileOutputStream("./db2.xlsx");
                workbook.write(outputStream);
                workbook.close();
                outputStream.close();

                continue;
            }


            if (companies.containsKey(companyName)) {
                System.out.println("InMemoryCache");
                Company company = companies.get(companyName);

                Row writeRow = sheet.getRow(i);
                Cell writeICell = writeRow.getCell(colIdxI, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                writeICell.setCellValue(company.getClassification());
                System.out.println("기업 구분 = " + company.getClassification());

                Cell writeHCell = writeRow.getCell(colIdxH, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                writeHCell.setCellValue(company.getBusiness());
                System.out.println("주요 사업 = " + company.getBusiness());

                FileOutputStream outputStream = new FileOutputStream("./db2.xlsx");
                workbook.write(outputStream);
                workbook.close();
                outputStream.close();

                continue;
            }

            boolean businessCheck = false, classificationCheck = false;

            if (!(companyName == null || companyName.equals(""))) {
                try {

                    String url = "https://www.jobkorea.co.kr/search/?stext=" + companyName + "&tabType=corp&Page_No=1";
                    Document docToCompanyUrl = Jsoup.connect(url).get();

                    Thread.sleep(3000);

                    Elements elements = null;
                    boolean tempCheck = false;
                    String tempName = null;

                    String[] elementsNames = new String[10];
                    Elements[] elementsValues = new Elements[10];
                    for (int j = 1; j < 10; j++) {
;                        Elements elementsName = docToCompanyUrl.select("#content > div > div > div.cnt-list-wrap > div > div.corp-info > div.lists > div > div.list-default > ul > li:nth-child(" + j + ") > div > div.post-list-corp.clear > div > a > strong");

                        if (!elementsName.isEmpty()) {
                            elements = docToCompanyUrl.select("#content > div > div > div.cnt-list-wrap > div > div.corp-info > div.lists > div > div.list-default > ul > li:nth-child(" + j + ") > div > div.post-list-corp.clear > div > a");


                            if (elements.get(0).attr("title").equals(companyName)) {
                                tempCheck = true;
                                break;
                            }

                            if (elements.get(0).attr("title") != null && !elements.get(0).attr("title").equals("")) {
                                tempName =elements.get(0).attr("title");
                            }

                            elementsNames[j] = elements.get(0).attr("title");
                            elementsValues[j] = elements;
                        }
                    }

                    if (elements != null && !tempCheck && tempName != null) {

                        for (int k = 1; k < 10; k++) {

                            if (elementsNames[k] == null || elementsValues[k] == null) {
                                continue;
                            }

                            try {
                                if (
                                        Math.abs(companyName.length() - tempName.length()) > Math.abs(companyName.length() - elementsNames[k].length()) &&
                                                elementsNames[k].contains(companyName)
                                ) {
                                    elements = elementsValues[k];
                                    tempName = elementsNames[k];
                                    System.out.println("tempName update= " + tempName);
                                }
                            } catch (Exception e) {
                                System.out.println("e = " + e);
                            }
                        }
                    }

                    if (elements == null) {
                        elements = docToCompanyUrl.select("#content > div > div > div.cnt-list-wrap > div > div.corp-info > div.lists > div > div.list-default > ul > li:nth-child(1) > div > div.post-list-corp.clear > div > a");
                    }

                    url = null;
                    for (Element element : elements) {
                        url = element.attr("href");
                        System.out.println("href url = " + url);
                    }

                    if (url == null) continue;

                    Document docToValues = Jsoup.connect("https://www.jobkorea.co.kr" + url).get();
                    Elements fields = docToValues.select("#company-body > div.company-body-infomation > div.company-infomation-row.basic-infomation > div > table > tbody");

                    for (Element field : fields) {
                        int count = 0;
                        Elements select = field.select("th.field-label");
                        Elements value = field.select("td.field-value div.value-container div.value");

                        for (Element label : select) {
                            if (businessCheck && classificationCheck) break;

                            String labelText = label.text();
                            if (labelText.equals("주요사업")) {
                                business = value.get(count).text();
                                if (value.get(count).text().startsWith("국")) {
                                    business = value.get(count - 1).text();
                                }
                                businessCheck = true;
                            }

                            else if (labelText.equals("기업구분")) {

                                if (
                                        companyMenu.equals("공공 기관") ||
                                        companyMenu.equals("행정 기관") ||
                                        companyMenu.equals("공공기관") ||
                                        companyMenu.equals("행정기관")
                                ) {

                                    if (!(
                                            value.get(count).text().equalsIgnoreCase("공공기관") ||
                                            value.get(count).text().equalsIgnoreCase("공공 기관") ||
                                            value.get(count).text().equalsIgnoreCase("국내 공공기관·공기업") ||
                                            value.get(count).text().equalsIgnoreCase("비영리법인") ||
                                                    value.get(count).text().contains("계열사")
                                    ) && !companyName.contains("은행")) {
                                        classification = "공공 기관";
                                        business = "공공 행정 업무";
                                        businessCheck = true;
                                        classificationCheck = true;
                                        break;
                                    }
                                }
                                classification = value.get(count).text();
                                classificationCheck = true;
                            }

                            if (businessCheck && classificationCheck) {
                                break;
                            }
                            count++;
                        }
                    }

                    Row writeRow = sheet.getRow(i);

                    System.out.println("기업명 = " + companyName);
                    if (classificationCheck) {
                        Cell writeICell = writeRow.getCell(colIdxI, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        writeICell.setCellValue(classification);
                        System.out.println("기업 구분 = " + classification);
                    }

                    if (businessCheck) {
                        Cell writeHCell = writeRow.getCell(colIdxH, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        writeHCell.setCellValue(business);
                        System.out.println("주요 사업 = " + business);
                    }

                    if (businessCheck && classificationCheck) {
                        companies.put(companyName, new Company(business, classification));
                    }

                } catch (Exception e) {
                    e.printStackTrace();

                } finally {
                    FileOutputStream outputStream = new FileOutputStream("./db2.xlsx");
                    workbook.write(outputStream);
                    workbook.close();
                    outputStream.close();
                }

            }
        }
    }
}
