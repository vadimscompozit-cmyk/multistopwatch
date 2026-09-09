package com.sasin.multistopwatch;

import android.annotation.SuppressLint;
import android.content.ContentResolver;
import android.content.ContentValues;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.provider.MediaStore;
import android.util.Base64;
import android.view.WindowManager;
import android.webkit.JavascriptInterface;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceResponse;
import android.webkit.WebSettings;
import android.webkit.WebView;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.FileProvider;
import androidx.webkit.WebSettingsCompat;
import androidx.webkit.WebViewAssetLoader;
import androidx.webkit.WebViewClientCompat;
import androidx.webkit.WebViewFeature;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Приложение целиком живёт в WebView: страница и её ресурсы лежат в assets,
 * замеры — в localStorage этого WebView. Java-часть нужна только там, где
 * веб-страница упирается в систему: файлы, «поделиться», экран, кнопка «назад».
 */
public class MainActivity extends AppCompatActivity {

    /** Ассеты отдаются по https-адресу: так у страницы полноценный
     *  безопасный источник, и localStorage не считается временным. */
    private static final String APP_URL =
            "https://appassets.androidplatform.net/assets/index.html";

    private WebView web;

    @SuppressLint("SetJavaScriptEnabled")
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        web = new WebView(this);
        setContentView(web);

        WebSettings settings = web.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true); // здесь хранятся секундомеры и архив
        settings.setTextZoom(100);           // системный масштаб шрифта не ломает вёрстку

        // Тёмная тема страницы включается вслед за системной.
        if (WebViewFeature.isFeatureSupported(WebViewFeature.ALGORITHMIC_DARKENING)) {
            WebSettingsCompat.setAlgorithmicDarkeningAllowed(settings, true);
        }

        final WebViewAssetLoader loader = new WebViewAssetLoader.Builder()
                .addPathHandler("/assets/", new WebViewAssetLoader.AssetsPathHandler(this))
                .build();

        web.setWebViewClient(new WebViewClientCompat() {
            @Override
            public WebResourceResponse shouldInterceptRequest(WebView view, WebResourceRequest request) {
                return loader.shouldInterceptRequest(request.getUrl());
            }
        });

        web.addJavascriptInterface(new Bridge(), "AndroidHost");
        web.loadUrl(APP_URL);
    }

    /** Кнопка «назад» сначала закрывает открытый лист внутри страницы. */
    @Override
    public void onBackPressed() {
        web.evaluateJavascript(
                "window.__androidBack ? window.__androidBack() : false",
                value -> {
                    if (!"true".equals(value)) finish();
                });
    }

    @Override
    protected void onDestroy() {
        if (web != null) {
            web.destroy();
            web = null;
        }
        super.onDestroy();
    }

    /** Методы этого класса видны странице как объект AndroidHost. */
    public class Bridge {

        /** Экран не гаснет, пока идёт хотя бы один секундомер. */
        @JavascriptInterface
        public void keepAwake(final boolean on) {
            runOnUiThread(() -> {
                if (on) {
                    getWindow().addFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);
                } else {
                    getWindow().clearFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);
                }
            });
        }

        /**
         * Кладёт файл в общую папку «Загрузки». Возвращает "ok" либо текст
         * ошибки — страница показывает его пользователю.
         */
        @JavascriptInterface
        public String saveFile(String name, String base64, String mime) {
            Uri uri = null;
            ContentResolver resolver = getContentResolver();
            try {
                byte[] bytes = Base64.decode(base64, Base64.DEFAULT);

                ContentValues values = new ContentValues();
                values.put(MediaStore.Downloads.DISPLAY_NAME, name);
                values.put(MediaStore.Downloads.MIME_TYPE, mime);
                values.put(MediaStore.Downloads.IS_PENDING, 1);

                uri = resolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values);
                if (uri == null) return "Не удалось создать файл в «Загрузках»";

                try (OutputStream out = resolver.openOutputStream(uri)) {
                    if (out == null) return "Не удалось открыть файл для записи";
                    out.write(bytes);
                }

                values.clear();
                values.put(MediaStore.Downloads.IS_PENDING, 0);
                resolver.update(uri, values, null, null);
                return "ok";
            } catch (Exception e) {
                if (uri != null) {
                    // недописанный файл не должен остаться висеть в «Загрузках»
                    try { resolver.delete(uri, null, null); } catch (Exception ignored) { }
                }
                return "Не удалось сохранить: " + e.getMessage();
            }
        }

        /** Отдаёт файл системному меню «Поделиться». */
        @JavascriptInterface
        public void shareFile(final String name, String base64, final String mime) {
            try {
                byte[] bytes = Base64.decode(base64, Base64.DEFAULT);

                File dir = new File(getCacheDir(), "exports");
                if (!dir.isDirectory() && !dir.mkdirs()) return;

                File file = new File(dir, name);
                try (FileOutputStream out = new FileOutputStream(file)) {
                    out.write(bytes);
                }

                final Uri uri = FileProvider.getUriForFile(
                        MainActivity.this, getPackageName() + ".files", file);

                runOnUiThread(() -> {
                    Intent send = new Intent(Intent.ACTION_SEND);
                    send.setType(mime);
                    send.putExtra(Intent.EXTRA_STREAM, uri);
                    send.putExtra(Intent.EXTRA_SUBJECT, name);
                    send.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                    startActivity(Intent.createChooser(send, getString(R.string.share_title)));
                });
            } catch (Exception ignored) {
                // Меню просто не откроется; файл всегда можно сохранить кнопкой рядом.
            }
        }
    }
}
