package com.poi.demo;

import android.os.Bundle;
import android.widget.FrameLayout;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;

import com.poi.demo.R;
import com.poi.office.PoiViewer;

public class FileViewActivity extends AppCompatActivity {

    private PoiViewer poiViewer;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_file_view);
        initView();
    }

    private void initView() {
        FrameLayout frameLayout = findViewById(R.id.layout_office);
        String filePath = getIntent().getStringExtra("filePath");
        poiViewer = new PoiViewer(this);
        try {
            poiViewer.loadFile(frameLayout, filePath);
        } catch (Exception e) {
            Toast.makeText(this, "打开失败", Toast.LENGTH_SHORT).show();
        }
    }

    @Override
    protected void onDestroy() {
        if (poiViewer != null) {
            poiViewer.recycle();
            poiViewer = null;
        }
        super.onDestroy();
    }
}