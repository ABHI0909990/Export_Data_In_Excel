package com.example.myapplication;

import android.Manifest;
import android.content.ActivityNotFoundException;
import android.content.ContentResolver;
import android.content.ContentValues;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
import android.widget.Button;
import android.widget.Toast;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import androidx.core.content.FileProvider;
import androidx.recyclerview.widget.LinearLayoutManager;
import androidx.recyclerview.widget.RecyclerView;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class MainActivity extends AppCompatActivity {

    private List<Customer> customerList;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        RecyclerView recyclerView = findViewById(R.id.recyclerView);
        Button btnExport = findViewById(R.id.btnExport);

        customerList = new ArrayList<>();
        customerList.add(new Customer("Void", "void@example.com", "1234567890"));
        customerList.add(new Customer("Michel", "michel@example.com", "0987654321"));
        customerList.add(new Customer("Sherif", "sherif@example.com", "8975629016"));
        customerList.add(new Customer("Steve", "Steve@example.com", "9064517830"));
        customerList.add(new Customer("Dustin", "dustin@example.com", "1348902645"));

        recyclerView.setLayoutManager(new LinearLayoutManager(this));
        recyclerView.setAdapter(new CustomerAdapter(customerList));

        btnExport.setOnClickListener(v -> exportToExcel());
        btnExport.setOnClickListener(v -> {
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                saveFileScopedStorage();
            } else {
                exportToExcel();
            }
        });

    }

    private void exportToExcel() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Customers");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Email");
        header.createCell(2).setCellValue("Phone");

        for (int i = 0; i < customerList.size(); i++) {
            Row row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(customerList.get(i).getName());
            row.createCell(1).setCellValue(customerList.get(i).getEmail());
            row.createCell(2).setCellValue(customerList.get(i).getPhone());
        }

        try {
            File downloadsDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS);
            if (!downloadsDir.exists()) {
                downloadsDir.mkdirs();
            }

            File file = new File(downloadsDir, "Customer.xlsx");
            FileOutputStream fileOut = new FileOutputStream(file);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            Toast.makeText(this, "File saved to: " + file.getAbsolutePath(), Toast.LENGTH_LONG).show();
            openFile(file);

        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Error exporting file: " + e.getMessage(), Toast.LENGTH_SHORT).show();
        }
    }


    private void checkPermissions() {
        if (Build.VERSION.SDK_INT <= Build.VERSION_CODES.P) { // For API 28 and below
            if (ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE)
                    != PackageManager.PERMISSION_GRANTED) {
                ActivityCompat.requestPermissions(this,
                        new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, 100);
            }
        }
    }
    private File getExportPath(String fileName) {
        File exportDir = new File(Environment.getExternalStorageDirectory(), "MyAppExports");
        if (!exportDir.exists()) {
            exportDir.mkdirs();
        }
        return new File(exportDir, fileName);
    }
    private void saveFileScopedStorage() {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
            ContentResolver resolver = getContentResolver();
            ContentValues contentValues = new ContentValues();
            contentValues.put(MediaStore.MediaColumns.DISPLAY_NAME, "Customer.xlsx");
            contentValues.put(MediaStore.MediaColumns.MIME_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            contentValues.put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOCUMENTS + "/MyAppExports");

            Uri uri = resolver.insert(MediaStore.Files.getContentUri("external"), contentValues);
            try {
                OutputStream outputStream = resolver.openOutputStream(uri);
                Workbook workbook = new XSSFWorkbook(); // Your Apache POI workbook
                Sheet sheet = workbook.createSheet("Customers");

                workbook.write(outputStream);
                workbook.close();
                outputStream.close();

                Toast.makeText(this, "File saved successfully", Toast.LENGTH_SHORT).show();
            } catch (Exception e) {
                e.printStackTrace();
                Toast.makeText(this, "Error: " + e.getMessage(), Toast.LENGTH_SHORT).show();
            }
        }
    }
    private void openFile(File file) {
        Intent intent = new Intent(Intent.ACTION_VIEW);
        Uri fileUri = FileProvider.getUriForFile(this, getPackageName() + ".fileprovider", file);
        intent.setDataAndType(fileUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        try {
            startActivity(intent);
        } catch (ActivityNotFoundException e) {
            Toast.makeText(this, "No app available to open Excel files", Toast.LENGTH_SHORT).show();
        }
    }
}
