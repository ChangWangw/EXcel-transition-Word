package com.example.poitest;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class UploadExcel {
    @RequestMapping("/uploadExcel")
    public boolean uploadExcel(@RequestParam MultipartFile file, HttpServletRequest request) throws IOException {

        if(!file.isEmpty()){
            String filePath = file.getOriginalFilename();
            //windows
            String savePath = request.getSession().getServletContext().getRealPath(filePath);

            //linux
            //String savePath = "/home/odcuser/webapps/file";

            File targetFile = new File(savePath);

            if(!targetFile.exists()){
                targetFile.mkdirs();
            }

            file.transferTo(targetFile);
            return true;
        }

        return false;
    }


    public static void readExcel() throws Exception{
        String fileName = "F:\\ChangWang\\上海市单位相关信息.xlsx";
        InputStream is = new FileInputStream(new File(fileName));
        Workbook hssfWorkbook = null;
        if (fileName.endsWith("xlsx")){
            hssfWorkbook = new XSSFWorkbook(is);//Excel 2007
        }else if (fileName.endsWith("xls")){
            hssfWorkbook = new HSSFWorkbook(is);//Excel 2003
        }
        List<ExcelBean> list = new ArrayList<ExcelBean>();
        // 循环工作表Sheet
        for (int numSheet = 0; numSheet <hssfWorkbook.getNumberOfSheets(); numSheet++) {
            Sheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
            if (hssfSheet == null) {
                continue;
            }
            for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                Row hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow != null) {
                    ExcelBean excelBean = new ExcelBean(hssfRow.getCell(0).toString(),hssfRow.getCell(1).toString(),hssfRow.getCell(2).toString().split("\\.")[0]);
                    System.out.println(excelBean.toString());
                    list.add(excelBean);
                }
            }
        }
        write2Docx(list);
    }

    public static void main(String[] args) throws Exception{
        readExcel();
    }


    public static void write2Docx(List<ExcelBean> list)throws Exception{
        XWPFDocument document= new XWPFDocument();
        File file = new File("D:\\Offer\\邮件3.docx");
        if(file.getParentFile().exists()) {//存放fileName文件的父目录存在
            try {
                file.createNewFile();//则创建fileName这个文件
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }else {//如果存放fileName文件的父目录不存在
            file.mkdirs(); //则创建整个父目录E:\eclipse-workspace\TheService1\src
            try {
                file.createNewFile();//并且在已创建好的父目录底下创建这个fileName文件
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
        FileOutputStream out = new FileOutputStream(file);
        //段落
        XWPFParagraph firstParagraph = document.createParagraph();
        firstParagraph.setAlignment(ParagraphAlignment.LEFT);
        //设置段落居中
        XWPFRun run = firstParagraph.createRun();
        int a = 0;
        for (ExcelBean excel : list) {
            run.setText("邮编：" + excel.getTrs3());
            run.addCarriageReturn();
            run.setText("" + excel.getTrs1());
            run.addCarriageReturn();
            run.setText("" + excel.getTrs2());
            run.addCarriageReturn();
            run.setText("办公室车管负责人：收");
            run.addCarriageReturn();
            if ((a+1)%4!=0) {
                run.addCarriageReturn();
                run.addCarriageReturn();
            }
            a++;
        }
        run.setColor("000000");//字体颜色
        run.setFontSize(16);//字体大小
        document.write(out);
        out.close();
    }
}
