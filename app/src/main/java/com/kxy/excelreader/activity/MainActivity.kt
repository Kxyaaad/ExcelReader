package com.kxy.excelreader.activity

import android.os.Bundle
import androidx.appcompat.app.AppCompatActivity
import com.kxy.excelreader.databinding.ActivityMainBinding

class MainActivity: AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(ActivityMainBinding.inflate(layoutInflater).root)
    }
}