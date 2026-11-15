package ar.gob.comuna4.relevamiento;

import android.content.ContentValues;
import android.content.Intent;
import android.net.Uri;
import android.os.Build;
import android.os.Environment;
import android.provider.MediaStore;
import android.util.Base64;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

import java.io.OutputStream;

@CapacitorPlugin(name = "Downloads")
public class DownloadsPlugin extends Plugin {
    @PluginMethod
    public void openPdf(PluginCall call) {
        String value = call.getString("uri");
        if (value == null || value.isEmpty()) {
            call.reject("No se indicó la ubicación del PDF.");
            return;
        }
        try {
            Intent view = new Intent(Intent.ACTION_VIEW);
            view.setDataAndType(Uri.parse(value), "application/pdf");
            view.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
            getActivity().startActivity(Intent.createChooser(view, "Abrir PDF con"));
            call.resolve();
        } catch (Exception error) {
            call.reject("No se encontró una aplicación para abrir el PDF: " + error.getMessage(), error);
        }
    }

    @PluginMethod
    public void savePdf(PluginCall call) {
        String data = call.getString("data");
        String fileName = call.getString("fileName");
        if (fileName == null || fileName.isEmpty()) fileName = "Acta de inspeccion.pdf";
        if (data == null || data.isEmpty()) {
            call.reject("No se recibieron los datos del PDF.");
            return;
        }
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.Q) {
            call.reject("El guardado directo en Descargas requiere Android 10 o superior.");
            return;
        }

        Uri uri = null;
        try {
            byte[] pdfBytes = Base64.decode(data, Base64.DEFAULT);
            ContentValues values = new ContentValues();
            values.put(MediaStore.Downloads.DISPLAY_NAME, fileName);
            values.put(MediaStore.Downloads.MIME_TYPE, "application/pdf");
            values.put(MediaStore.Downloads.RELATIVE_PATH,
                Environment.DIRECTORY_DOWNLOADS + "/Comuna 4 - Relevamientos");
            values.put(MediaStore.Downloads.IS_PENDING, 1);

            uri = getContext().getContentResolver().insert(
                MediaStore.Downloads.EXTERNAL_CONTENT_URI, values);
            if (uri == null) throw new Exception("Android no pudo crear el archivo en Descargas.");

            try (OutputStream output = getContext().getContentResolver().openOutputStream(uri)) {
                if (output == null) throw new Exception("No se pudo abrir el archivo de destino.");
                output.write(pdfBytes);
                output.flush();
            }
            ContentValues published = new ContentValues();
            published.put(MediaStore.Downloads.IS_PENDING, 0);
            getContext().getContentResolver().update(uri, published, null, null);

            JSObject result = new JSObject();
            result.put("uri", uri.toString());
            result.put("fileName", fileName);
            result.put("folder", "Descargas/Comuna 4 - Relevamientos");
            call.resolve(result);
        } catch (Exception error) {
            if (uri != null) getContext().getContentResolver().delete(uri, null, null);
            call.reject("No se pudo guardar el PDF en Descargas: " + error.getMessage(), error);
        }
    }
}
