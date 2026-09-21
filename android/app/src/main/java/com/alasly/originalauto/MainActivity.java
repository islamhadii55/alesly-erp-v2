package com.alasly.originalauto;

import android.app.AlertDialog;
import android.content.SharedPreferences;
import android.os.Bundle;
import android.webkit.ServiceWorkerClient;
import android.webkit.ServiceWorkerController;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceResponse;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.EditText;
import androidx.appcompat.app.AppCompatActivity;

public class MainActivity extends AppCompatActivity {
    private WebView web;
    private SharedPreferences prefs;
    private static final String PREFS = "alasly";
    private static final String KEY_URL = "server_url";
    private static final String DEFAULT_URL = "http://10.0.2.2:5000";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        prefs = getSharedPreferences(PREFS, MODE_PRIVATE);
        web = findViewById(R.id.web);

        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setDatabaseEnabled(true);
        s.setCacheMode(WebSettings.LOAD_DEFAULT);
        s.setAllowFileAccess(true);
        s.setMixedContentMode(WebSettings.MIXED_CONTENT_ALWAYS_ALLOW);

        if (android.os.Build.VERSION.SDK_INT >= android.os.Build.VERSION_CODES.N) {
            ServiceWorkerController.getInstance().setServiceWorkerClient(new ServiceWorkerClient() {
                @Override
                public WebResourceResponse shouldInterceptRequest(WebResourceRequest request) {
                    return null;
                }
            });
        }

        web.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
                return false;
            }
        });
        web.setWebChromeClient(new WebChromeClient());
        web.setOnLongClickListener(v -> {
            askServer();
            return true;
        });
        web.evaluateJavascript("window.ALASLY_PLATFORM='android';", null);
        web.loadUrl(prefs.getString(KEY_URL, DEFAULT_URL));
    }

    private void askServer() {
        final EditText input = new EditText(this);
        input.setText(prefs.getString(KEY_URL, DEFAULT_URL));
        new AlertDialog.Builder(this)
            .setTitle("عنوان سيرفر المحل")
            .setMessage("مثال على الشبكة المحلية: http://192.168.1.10:5000")
            .setView(input)
            .setPositiveButton("حفظ", (d, w) -> {
                String url = input.getText().toString().trim();
                prefs.edit().putString(KEY_URL, url).apply();
                web.loadUrl(url);
            })
            .setNegativeButton("إلغاء", null)
            .show();
    }

    @Override
    public void onBackPressed() {
        if (web.canGoBack()) web.goBack();
        else super.onBackPressed();
    }
}
