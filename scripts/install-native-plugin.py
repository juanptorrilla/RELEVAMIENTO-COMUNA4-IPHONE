from pathlib import Path
import re

android = Path('android')
main_files = list(android.rglob('MainActivity.java'))
if len(main_files) != 1:
    raise SystemExit(f'Se esperaba un MainActivity.java y se encontraron {len(main_files)}')
main = main_files[0]

for plugin_name in ('DownloadsPlugin', 'PhotoPickerPlugin'):
    source = Path('native') / f'{plugin_name}.java'
    target = main.parent / f'{plugin_name}.java'
    target.write_text(source.read_text(encoding='utf-8'), encoding='utf-8')

text = main.read_text(encoding='utf-8')
if 'import android.os.Bundle;' not in text:
    text = text.replace('import com.getcapacitor.BridgeActivity;',
                        'import com.getcapacitor.BridgeActivity;\nimport android.os.Bundle;')

pattern = r'public class MainActivity extends BridgeActivity\s*\{'
if 'registerPlugin(DownloadsPlugin.class)' not in text:
    replacement = '''public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(Bundle savedInstanceState) {
        registerPlugin(DownloadsPlugin.class);
        registerPlugin(PhotoPickerPlugin.class);
        super.onCreate(savedInstanceState);
    }
'''
    text, count = re.subn(pattern, lambda m: replacement, text, count=1)
    if count != 1:
        raise SystemExit('No se pudo registrar los plugins en MainActivity.java')
elif 'registerPlugin(PhotoPickerPlugin.class)' not in text:
    text = text.replace('registerPlugin(DownloadsPlugin.class);',
                        'registerPlugin(DownloadsPlugin.class);\n        registerPlugin(PhotoPickerPlugin.class);', 1)

main.write_text(text, encoding='utf-8')
print(f'Plugins nativos instalados y registrados en {main}')
