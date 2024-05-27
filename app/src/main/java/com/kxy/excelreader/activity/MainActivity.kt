package com.kxy.excelreader.activity

import android.Manifest
import android.annotation.SuppressLint
import android.content.pm.PackageManager
import android.os.Bundle
import android.util.Log
import android.webkit.WebSettings
import android.webkit.WebView
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import com.kxy.excelreader.R
import com.kxy.excelreader.databinding.ActivityMainBinding
import java.io.File

import com.kxy.officereader.ExcelReader
import com.kxy.officereader.WordReader
import java.io.BufferedWriter
import java.io.FileWriter
import java.io.IOException


class MainActivity : AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(ActivityMainBinding.inflate(layoutInflater).root)
//        getFile()
        val htmlContent = WordReader.readWord("/storage/emulated/0/Download/东部经开区智慧政务平台-经开通APP-用户操作手册-3.docx")

        val outputFilePath = "/storage/emulated/0/Download/东部经开区智慧政务平台-经开通APP-用户操作手册.html"
        try  {
            BufferedWriter(FileWriter(outputFilePath)).use { writer ->
                writer.write(htmlContent)
            }
        }catch (e: IOException) {
            Log.e("写入错误", e.message ?: "未知错误")
        }
        val webView = findViewById<WebView>(R.id.webView)
        val setting = webView.settings
        setting.javaScriptEnabled = true
        setting.builtInZoomControls = true
        setting.cacheMode = WebSettings.LOAD_CACHE_ELSE_NETWORK
//            webView.setInitialScale(300)
       println("html内容 => $htmlContent")
//        webView.loadData(htmlContent.toString(), "text/html", "utf-8")
        webView.loadUrl(outputFilePath)
    }

    private fun wordToPdf() {

    }

    @SuppressLint("SetJavaScriptEnabled")
    private fun getFile() {
        if (ContextCompat.checkSelfPermission(
                this,
                Manifest.permission.READ_EXTERNAL_STORAGE
            ) != PackageManager.PERMISSION_GRANTED
        ) {
            // 如果没有权限，则请求权限
            val REQUEST_CODE_STORAGE_PERMISSION = 10086
            ActivityCompat.requestPermissions(
                this,
                arrayOf(
                    Manifest.permission.WRITE_EXTERNAL_STORAGE,
                    Manifest.permission.READ_EXTERNAL_STORAGE
                ),
                REQUEST_CODE_STORAGE_PERMISSION
            )
        }
        val file = File(
            "/storage/emulated/0/Android/data/com.wb.jkt/files/ZONE_XBPM/process/四川省工程技术人员职称申报评审基本条件.xlsx"
        )

        Log.e("打开文件", file.absolutePath)
        val htmlContent = ExcelReader().readExcel(file.absolutePath)
        val webView = findViewById<WebView>(R.id.webView)
        val setting = webView.settings
        setting.javaScriptEnabled = true
        setting.builtInZoomControls = true
        setting.cacheMode = WebSettings.LOAD_CACHE_ELSE_NETWORK
//            webView.setInitialScale(300)
        Log.e("HTML内容", htmlContent.toString())
        webView.loadData(htmlContent.toString(), "text/html", "utf-8")

    }

}

