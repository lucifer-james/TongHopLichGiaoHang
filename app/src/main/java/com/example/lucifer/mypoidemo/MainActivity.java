package com.example.lucifer.mypoidemo;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.fragment.app.Fragment;

import android.os.Bundle;
import android.os.StrictMode;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.ProgressBar;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {

    Button btnTongHopLichGiaohang;
    Button btnShareFileTo;
    TextView txtResult;
    XSSFSheet sheet_goc;
    ProgressBar progressBar;

    String[] filePath = new String[4];

    static{
        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory","com.fasterxml.aalto.stax.InputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory","com.fasterxml.aalto.stax.OnputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory","com.fasterxml.aalto.stax.EventFactoryImpl");
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        initView();
    }

    private void initView() {
        btnTongHopLichGiaohang = (Button) findViewById(R.id.buttonTongHopLichGiaoHang);
        btnTongHopLichGiaohang.setOnClickListener(this);

        btnShareFileTo = (Button) findViewById(R.id.buttonShareFileTo);
        btnShareFileTo.setOnClickListener(this);

        txtResult = (TextView)findViewById(R.id.textViewResult);
        txtResult.setVisibility(View.INVISIBLE);

        progressBar = (ProgressBar)findViewById(R.id.progressbarXuLy);
        progressBar.setVisibility(View.INVISIBLE);

        ActivityCompat.requestPermissions(this,new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE,
            Manifest.permission.READ_EXTERNAL_STORAGE}, PackageManager.PERMISSION_GRANTED);

        //d??ng ????? gi???i quy???t vi???c Uri.parse("file:///storage/emulated/0/Download/KH GIAO BTPND NOI DIA 05-08-2021.xlsx")
        StrictMode.VmPolicy.Builder builder = new StrictMode.VmPolicy.Builder();
        StrictMode.setVmPolicy(builder.build());

    }

    @Override
    public void onClick(View view) {
        switch (view.getId()){
            case R.id.buttonTongHopLichGiaoHang:
                btnTongHopLichGiaohang_click();
                break;
            case R.id.buttonShareFileTo:
                btnShareFileTo_click();
                break;
            default: break;
        }

    }

    private void btnShareFileTo_click() {
        List<Fragment> list_uri = (List<Fragment>) getSupportFragmentManager().getFragments();

        String filePath="";
        fileChooserFragment frgFileChooser = (fileChooserFragment) list_uri.get(0);
        filePath = frgFileChooser.lbFilePath.getText().toString();

        File fileSend = new File(filePath);

        if(!fileSend.exists()){
            Toast toast = Toast.makeText(getApplicationContext(),"File kh??ng t???n t???i",Toast.LENGTH_SHORT);
            return;
        }

        Intent shareFile = new Intent();
        shareFile.setAction(Intent.ACTION_SEND);
        shareFile.setType("application/excel");
        shareFile.putExtra(Intent.EXTRA_STREAM, Uri.parse("file://"+fileSend));

        Intent chooser = Intent.createChooser(shareFile,"Share file to :");

        startActivity(chooser);

    }

    private void btnTongHopLichGiaohang_click(){
        txtResult.setVisibility(View.VISIBLE);
        txtResult.setText("??ang x??? l?? ......");

        int i,j;

        List<Fragment> listFragmentFile = (List<Fragment>) getSupportFragmentManager().getFragments();

        List<XSSFWorkbook> list_work_book = new ArrayList<XSSFWorkbook>();
        List<XSSFSheet> list_work_sheet = new ArrayList<XSSFSheet>();

        Object cellData = null;

        String str_so_tong = "";

        try {
            String result ="";

            for (i = 0; i < listFragmentFile.size(); i++) {
                fileChooserFragment frgFileChooser = (fileChooserFragment) listFragmentFile.get(i);
                filePath[i] = frgFileChooser.lbFilePath.getText().toString();

                File file = new File(filePath[i]);
                FileInputStream fileInputStream = new FileInputStream(file);

                XSSFWorkbook wbtemp = new XSSFWorkbook(fileInputStream);
                XSSFSheet wstemp = wbtemp.getSheetAt(0);

                list_work_book.add( wbtemp);
                list_work_sheet.add(wstemp);

                fileInputStream.close();
            }

            txtResult.setText("???? l???y ???????c file");

            int end_row = list_work_sheet.get(0).getLastRowNum();

            //int end_row = 38;

            //d??ng b???t ?????u c?? m?? BTP
            int row = 11;

            list_work_book.get(0).setForceFormulaRecalculation(true);

            // l???y ra sheet g???c
            sheet_goc = list_work_sheet.get(0);

            //ch???y th??? xem thread c?? x??? l?? d?????c kh??ng ?
            ThreadXuLyFile threadXuLyFile = new ThreadXuLyFile(list_work_sheet);
            threadXuLyFile.start();
            try {
                // join ngh??a l?? khi thread n??y ch???y xong th?? c??c thread kh??c hay process kh??c m???i ???????c ch???y
                threadXuLyFile.join();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }

            int row_Ekitai = threadXuLyFile.getRowTypeBTP("BTP G??I EKITAI");

            txtResult.setText("D???n d???p xong file g???c");

            // ??o???n m?? x??? l?? ch??p d??? li???u v??o sheet t???ng h???p
            // duy???t t??? d??ng 0 ?????n d??ng cu???i c???a sheet
            while(row < end_row){
                try {
                    // l???y ra t???ng cell ch???a m?? BTP (c???t 2)
                    XSSFCell cell_goc = sheet_goc.getRow(row).getCell(2);

                    // l???y gi?? tr??? c???a cell ???? (m?? BTP)
                    Object cell_goc_value = getCellValue(cell_goc, list_work_book.get(0).getCreationHelper().createFormulaEvaluator());

                    // n???u ?? ???? c?? m?? th?? m???i x??? l??, kh??ng th?? b??? qua
                    if(!cell_goc_value.equals("") && !cell_goc_value.equals(0)){
                        // duy???t qua t???ng sheet c??n l???i (index = 1 to 3)
                        // index = 0 l?? sheet g???c
                        for (int k = 1; k < list_work_sheet.size(); k++) {

                            // bi???n row v?? end_row t???m ????? duy???t qua t???ng sheet
                            int row_temp = row;
                            int end_row_temp = end_row;

                            //duy???t t??? d??ng 11 ?????n h???t c???a t???ng sheet
                            while (row_temp < end_row_temp) {
                                int m=0;
                                try {
                                    // c??ng l???y ra m?? BTP c???a t???ng sheet
                                    // ?????t trong try catch NullPointerException ????? n???u cell = null th?? ti???p t???c l???y cell d?????i
                                    XSSFCell cell_temp = list_work_sheet.get(k).getRow(row_temp).getCell(2);
                                    Object cell_temp_value = getCellValue(cell_temp, list_work_book.get(k).getCreationHelper().createFormulaEvaluator());

                                    // n???u m?? 2 BTP b???ng nhau (m?? trong sheet g???c v?? m?? trong sheet c???n t???ng k???t)
                                    if (cell_temp_value.equals(cell_goc_value)) {
                                        // duy???t qua t???ng c???t
                                        // c???t b???t ?????u c?? BTP l?? c???t 18
                                        // c???t k???t th??c l???ch giao h??ng l?? c???t 52
                                        for (m = 18; m < 53; m++) {
                                            // l???y ra cell trong file g???c
                                            XSSFCell cell_data_goc = sheet_goc.getRow(row).getCell(m);
                                            cellData = getCellValue(cell_data_goc,list_work_book.get(0).getCreationHelper().createFormulaEvaluator());

                                            str_so_tong = cellData.toString();

                                            // n???u cell == null th?? t???o cell m???i
                                            if (cell_data_goc == null) {
                                                sheet_goc.getRow(row).createCell(m);
                                            }

                                            XSSFCell cell_data_person = list_work_sheet.get(k).getRow(row_temp).getCell(m);

                                            // so_tong l?? 1 Object ch???a gi?? tr??? c???a cell l???y ra
                                            cellData = getCellValue(cell_data_person, list_work_book.get(k).getCreationHelper().createFormulaEvaluator());

                                            if (cell_data_person.getCellComment()!=null){
                                                String comment = cell_data_person.getCellComment().getString().toString();
                                                try{
                                                    setCellComment(cell_data_goc,comment);
                                                }catch (IllegalArgumentException e){
                                                    e.printStackTrace();
                                                }
                                            }
                                            // chuy???n s??? t???ng ???? g??n v??o str_so_tong

                                            if (row <= row_Ekitai){
                                                str_so_tong += "+" + cellData.toString();
                                            }else{
                                                double temp = Double.parseDouble(cellData.toString())/3;
                                                Log.d("So tong chua cong :", str_so_tong);
                                                str_so_tong += "+" + String.valueOf(temp) + "-" + String.valueOf(temp) ;
                                                Log.d("Truy EKITAI:","Ekitai :" + cell_goc_value + " c?? gi?? tr??? l??: " +
                                                        cellData.toString() + " ??? d??ng " + row +" - c???t " +m );
                                                Log.d("So tong da cong:", str_so_tong);
                                            }

                                            // g??n c??ng th???c cho ?? trong sheet g???c
                                            cell_data_goc.setCellFormula(str_so_tong);
                                        }
                                    }
                                }catch (NullPointerException e){
                                    e.printStackTrace();
                                }catch (FormulaParseException e){
                                    e.printStackTrace();
                                    Log.d("L???i t???i:","Sheet s??? " + String.valueOf(k) + " - d??ng s??? " +
                                            String.valueOf(row_temp) + " - c???t s??? :" + String.valueOf(m) + " - gi?? tr??? : " + String.valueOf(cellData));
                                }
                                row_temp++;
                            }
                        }
                    }

                }catch(NullPointerException e){
                    e.printStackTrace();
                }

                row++;
            }

            File file = new File(filePath[0]);

            if(file.exists()){
                file.delete();
                file.createNewFile();
            }

            FileOutputStream fos = new FileOutputStream(file);
            list_work_book.get(0).write(fos);

            fos.close();
            list_work_book.get(0).close();

            Toast toast = Toast.makeText(getApplicationContext(),"???? ghi file thanh cong",Toast.LENGTH_SHORT);
            toast.show();

            txtResult.setText("???? ho??n th??nh !!!");


        }catch (ClassCastException e){
            e.printStackTrace();
        }catch (FileNotFoundException e){
            e.printStackTrace();
        }catch (NoClassDefFoundError e){
            e.printStackTrace();
        } catch (InvalidOperationException e){
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }catch (RuntimeException e) {
            e.printStackTrace();
        }
    }

    private Object getCellValue(XSSFCell cell, FormulaEvaluator evaluator){

        Object cellValue = evaluator.evaluate(cell);

        switch (cell.getCellTypeEnum()) {
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case BLANK:
                cellValue = 0;
                break;
            case _NONE:
                cellValue = 0;
                break;
            case ERROR:
                cellValue = 0;
                break;
            case FORMULA:
                cellValue = evaluator.evaluate(cell).getNumberValue();
                break;
            default:
                return 0;
        }

        return cellValue;
    }

    private void setCellComment(XSSFCell cell, String comment){
        Drawing drawing = cell.getSheet().createDrawingPatriarch();
        CreationHelper creationHelper = cell.getSheet().getWorkbook().getCreationHelper();
        ClientAnchor clientAnchor = creationHelper.createClientAnchor();

        clientAnchor.setCol1(cell.getColumnIndex());
        clientAnchor.setCol2(cell.getColumnIndex()+2);
        clientAnchor.setRow1(cell.getRowIndex());
        clientAnchor.setRow2(cell.getRowIndex()+4);

        clientAnchor.setDx1(300);
        clientAnchor.setDx2(300);
        clientAnchor.setDy1(300);
        clientAnchor.setDy2(300);

        Comment cmt = drawing.createCellComment(clientAnchor);
        RichTextString richTextString = creationHelper.createRichTextString(comment);

        cmt.setString(richTextString);

        cell.setCellComment(cmt);

    }

}
