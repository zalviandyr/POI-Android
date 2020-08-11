package com.zukron.poi;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import android.Manifest;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.os.Environment;
import android.text.TextUtils;
import android.view.View;
import android.widget.Button;
import android.widget.Toast;

import com.google.android.material.textfield.TextInputEditText;
import com.google.android.material.textfield.TextInputLayout;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {
    private TextInputLayout inputLayoutText;
    private TextInputEditText inputText;
    private Button btnCreateDocx, btnCreateXslx, btnCreatePptx;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl");

        inputLayoutText = findViewById(R.id.input_layout_text);
        inputText = findViewById(R.id.input_text);
        btnCreateDocx = findViewById(R.id.btn_create_docx);
        btnCreateXslx = findViewById(R.id.btn_create_xlsx);
        btnCreatePptx = findViewById(R.id.btn_create_pptx);
    }

    @Override
    protected void onStart() {
        super.onStart();

        if (checkSelfPermission(Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
            ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, PackageManager.PERMISSION_GRANTED);
        }

        btnCreateDocx.setOnClickListener(this);
        btnCreateXslx.setOnClickListener(this);
        btnCreatePptx.setOnClickListener(this);
    }

    @Override
    public void onClick(View view) {
        File path = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS);
        if (validate()) {
            assert inputText.getText() != null;
            String message = inputText.getText().toString().trim();

            switch (view.getId()) {
                case R.id.btn_create_docx:
                    createDocx(path, message);
                    break;
                case R.id.btn_create_xlsx:
                    createXslx(path, message);
                    break;
                case R.id.btn_create_pptx:
                    createPptx(path, message);
                    break;
            }
        }
    }

    private boolean validate() {
        boolean valid = true;

        assert inputText.getText() != null;
        if (TextUtils.isEmpty(inputText.getText().toString().trim())) {
            inputLayoutText.setError("Harus diisi");
            valid = false;
        }

        return valid;
    }

    private void createDocx(File path, String msg) {
        try {
            File file = new File(path, "/poi.docx");
            XWPFDocument xwpfDocument = new XWPFDocument();
            FileOutputStream fileOutputStream = new FileOutputStream(file);

            XWPFParagraph xwpfParagraph = xwpfDocument.createParagraph();
            XWPFRun xwpfRun = xwpfParagraph.createRun();
            xwpfRun.setText(msg);

            xwpfDocument.write(fileOutputStream);
            fileOutputStream.close();
            Toast.makeText(this, "poi.docx sukses dibuat", Toast.LENGTH_SHORT).show();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void createXslx(File path, String msg) {
        try {
            File file = new File(path, "/poi.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook();
            FileOutputStream fileOutputStream = new FileOutputStream(file);

            XSSFSheet sheet = workbook.createSheet();
            XSSFRow row = sheet.createRow(2);
            XSSFCell cell = row.createCell(1);
            cell.setCellValue(msg);

            workbook.write(fileOutputStream);
            fileOutputStream.close();
            Toast.makeText(this, "poi.xlsx sukses dibuat", Toast.LENGTH_SHORT).show();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void createPptx(File path, String msg) {
        try {
            File file = new File(path, "/pot.pptx");
            XMLSlideShow slideShow = new XMLSlideShow();
            FileOutputStream fileOutputStream = new FileOutputStream(file);

            XSLFSlideMaster slideMaster = slideShow.getSlideMasters().get(0);
            XSLFSlideLayout slideLayout = slideMaster.getLayout(SlideLayout.TITLE);
            XSLFSlide slide = slideShow.createSlide(slideLayout);
            XSLFTextShape title = slide.getPlaceholder(0);
            title.setText(msg);

            slideShow.write(fileOutputStream);
            fileOutputStream.close();
            Toast.makeText(this, "poi.pptx sukses dibuat", Toast.LENGTH_SHORT).show();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}