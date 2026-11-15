package ar.gob.comuna4.relevamiento;

import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.provider.OpenableColumns;
import android.util.Base64;

import androidx.activity.result.ActivityResult;
import com.getcapacitor.JSArray;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.ActivityCallback;
import com.getcapacitor.annotation.CapacitorPlugin;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

@CapacitorPlugin(name = "PhotoPicker")
public class PhotoPickerPlugin extends Plugin {
    @PluginMethod
    public void pickImages(PluginCall call) {
        Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        intent.setType("image/*");
        intent.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, true);
        startActivityForResult(call, intent, "handlePickerResult");
    }

    @ActivityCallback
    private void handlePickerResult(PluginCall call, ActivityResult result) {
        if (call == null) return;
        if (result.getResultCode() != android.app.Activity.RESULT_OK || result.getData() == null) {
            JSObject empty = new JSObject();
            empty.put("photos", new JSArray());
            call.resolve(empty);
            return;
        }
        try {
            Intent data = result.getData();
            JSArray photos = new JSArray();
            int max = 4;
            if (data.getClipData() != null) {
                int count = Math.min(data.getClipData().getItemCount(), max);
                for (int i = 0; i < count; i++) {
                    Uri uri = data.getClipData().getItemAt(i).getUri();
                    if (uri != null) photos.put(readPhoto(uri));
                }
            } else if (data.getData() != null) {
                photos.put(readPhoto(data.getData()));
            }
            JSObject response = new JSObject();
            response.put("photos", photos);
            call.resolve(response);
        } catch (Exception error) {
            call.reject("No se pudieron leer las fotografías seleccionadas: " + error.getMessage(), error);
        }
    }

    private JSObject readPhoto(Uri uri) throws Exception {
        String mime = getContext().getContentResolver().getType(uri);
        if (mime == null || !mime.startsWith("image/")) mime = "image/jpeg";
        String name = getFileName(uri);
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (InputStream input = getContext().getContentResolver().openInputStream(uri)) {
            if (input == null) throw new Exception("No se pudo abrir la fotografía.");
            byte[] buffer = new byte[8192];
            int read;
            while ((read = input.read(buffer)) != -1) output.write(buffer, 0, read);
        }
        JSObject photo = new JSObject();
        photo.put("name", name);
        photo.put("mimeType", mime);
        photo.put("data", Base64.encodeToString(output.toByteArray(), Base64.NO_WRAP));
        return photo;
    }

    private String getFileName(Uri uri) {
        String name = "fotografia.jpg";
        try (Cursor cursor = getContext().getContentResolver().query(uri, null, null, null, null)) {
            if (cursor != null && cursor.moveToFirst()) {
                int index = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME);
                if (index >= 0) name = cursor.getString(index);
            }
        } catch (Exception ignored) {}
        return name;
    }
}
