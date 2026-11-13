package org.opencells.app;

import android.app.Activity;
import android.content.ContentResolver;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.provider.OpenableColumns;
import android.util.Base64;
import androidx.activity.result.ActivityResult;
import androidx.core.content.FileProvider;
import com.getcapacitor.JSArray;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.ActivityCallback;
import com.getcapacitor.annotation.CapacitorPlugin;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * Storage Access Framework bridge.
 * open():   system picker -> copies the document to a private cache file the WebView can fetch()
 *           (no huge base64 strings for big spreadsheets) + keeps a persistable read/write grant.
 * write():  overwrite the original document in place ("Save").
 * create(): system "Save as" picker.
 * share():  share sheet via FileProvider.
 * No storage permissions are needed on any Android version (8.0+).
 */
@CapacitorPlugin(name = "Documents")
public class DocumentsPlugin extends Plugin {
    private final ExecutorService io = Executors.newSingleThreadExecutor();

    @PluginMethod
    public void open(PluginCall call) {
        Intent i = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        i.addCategory(Intent.CATEGORY_OPENABLE);
        i.setType("*/*");
        JSArray types = call.getArray("mimeTypes");
        if (types != null && types.length() > 0) {
            try {
                String[] arr = new String[types.length()];
                for (int k = 0; k < types.length(); k++) arr[k] = types.getString(k);
                i.putExtra(Intent.EXTRA_MIME_TYPES, arr);
            } catch (Exception ignored) {
            }
        }
        i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION | Intent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION);
        startActivityForResult(call, i, "onOpened");
    }

    @ActivityCallback
    private void onOpened(PluginCall call, ActivityResult result) {
        if (call == null) return;
        if (result.getResultCode() != Activity.RESULT_OK || result.getData() == null || result.getData().getData() == null) {
            call.reject("cancelled");
            return;
        }
        Uri uri = result.getData().getData();
        ContentResolver cr = getContext().getContentResolver();
        try {
            cr.takePersistableUriPermission(uri, Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
        } catch (SecurityException e) {
            try {
                cr.takePersistableUriPermission(uri, Intent.FLAG_GRANT_READ_URI_PERMISSION);
            } catch (SecurityException ignored) {
            }
        }
        copyToCache(call, uri);
    }

    @PluginMethod
    public void read(PluginCall call) {
        String u = call.getString("uri");
        if (u == null) {
            call.reject("uri required");
            return;
        }
        copyToCache(call, Uri.parse(u));
    }

    private void copyToCache(PluginCall call, Uri uri) {
        io.execute(() -> {
            try {
                String name = displayName(uri);
                File dir = new File(getContext().getCacheDir(), "open");
                dir.mkdirs();
                File[] old = dir.listFiles();
                if (old != null) for (File f : old) f.delete(); // keep only the current file
                File out = new File(dir, System.currentTimeMillis() + "_" + name.replaceAll("[^A-Za-z0-9._-]", "_"));
                long size = 0;
                try (InputStream in = getContext().getContentResolver().openInputStream(uri); OutputStream os = new FileOutputStream(out)) {
                    if (in == null) throw new Exception("Cannot open document");
                    byte[] buf = new byte[1 << 16];
                    int n;
                    while ((n = in.read(buf)) > 0) {
                        os.write(buf, 0, n);
                        size += n;
                    }
                }
                JSObject ret = new JSObject();
                ret.put("uri", uri.toString());
                ret.put("name", name);
                ret.put("size", size);
                ret.put("path", out.getAbsolutePath());
                call.resolve(ret);
            } catch (Exception e) {
                call.reject("Could not read file: " + e.getMessage(), e);
            }
        });
    }

    @PluginMethod
    public void write(PluginCall call) {
        String u = call.getString("uri");
        String data = call.getString("data");
        if (u == null || data == null) {
            call.reject("uri and data required");
            return;
        }
        io.execute(() -> {
            try (OutputStream os = getContext().getContentResolver().openOutputStream(Uri.parse(u), "wt")) {
                if (os == null) throw new Exception("Cannot write document");
                os.write(Base64.decode(data, Base64.DEFAULT));
                call.resolve();
            } catch (Exception e) {
                call.reject("Could not save: " + e.getMessage(), e);
            }
        });
    }

    @PluginMethod
    public void create(PluginCall call) {
        Intent i = new Intent(Intent.ACTION_CREATE_DOCUMENT);
        i.addCategory(Intent.CATEGORY_OPENABLE);
        i.setType(call.getString("mimeType", "application/octet-stream"));
        i.putExtra(Intent.EXTRA_TITLE, call.getString("name", "Untitled.xlsx"));
        startActivityForResult(call, i, "onCreated");
    }

    @ActivityCallback
    private void onCreated(PluginCall call, ActivityResult result) {
        if (call == null) return;
        if (result.getResultCode() != Activity.RESULT_OK || result.getData() == null || result.getData().getData() == null) {
            call.reject("cancelled");
            return;
        }
        Uri uri = result.getData().getData();
        try {
            getContext().getContentResolver().takePersistableUriPermission(uri, Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
        } catch (SecurityException ignored) {
        }
        String data = call.getString("data", "");
        io.execute(() -> {
            try (OutputStream os = getContext().getContentResolver().openOutputStream(uri, "wt")) {
                if (os == null) throw new Exception("Cannot write document");
                os.write(Base64.decode(data, Base64.DEFAULT));
                JSObject ret = new JSObject();
                ret.put("uri", uri.toString());
                ret.put("name", displayName(uri));
                call.resolve(ret);
            } catch (Exception e) {
                call.reject("Could not save: " + e.getMessage(), e);
            }
        });
    }

    @PluginMethod
    public void share(PluginCall call) {
        String name = call.getString("name", "workbook.xlsx");
        String mime = call.getString("mimeType", "application/octet-stream");
        String data = call.getString("data", "");
        io.execute(() -> {
            try {
                File dir = new File(getContext().getCacheDir(), "share");
                dir.mkdirs();
                File f = new File(dir, name.replaceAll("[/\\\\]", "_"));
                try (FileOutputStream os = new FileOutputStream(f)) {
                    os.write(Base64.decode(data, Base64.DEFAULT));
                }
                Uri uri = FileProvider.getUriForFile(getContext(), getContext().getPackageName() + ".fileprovider", f);
                Intent send = new Intent(Intent.ACTION_SEND);
                send.setType(mime);
                send.putExtra(Intent.EXTRA_STREAM, uri);
                send.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                getActivity().runOnUiThread(() -> {
                    getActivity().startActivity(Intent.createChooser(send, name));
                    call.resolve();
                });
            } catch (Exception e) {
                call.reject("Could not share: " + e.getMessage(), e);
            }
        });
    }

    private String displayName(Uri uri) {
        String name = null;
        try (Cursor c = getContext().getContentResolver().query(uri, new String[] { OpenableColumns.DISPLAY_NAME }, null, null, null)) {
            if (c != null && c.moveToFirst()) name = c.getString(0);
        } catch (Exception ignored) {
        }
        if (name == null) {
            String p = uri.getLastPathSegment();
            name = p != null ? p.substring(p.lastIndexOf('/') + 1) : "workbook.xlsx";
        }
        return name;
    }
}
