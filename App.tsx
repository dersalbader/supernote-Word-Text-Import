import React, {useState} from 'react';
import {Clipboard, Pressable, StyleSheet, Text, TouchableOpacity, View} from 'react-native';
import {PluginCommAPI, PluginFileAPI, PluginManager, PluginNoteAPI, RattaFileSelector} from 'sn-plugin-lib';
import {strFromU8, unzipSync} from 'fflate';

const extractText = async (filePath: string) => {
  const url = filePath.startsWith('file://') ? filePath : `file://${filePath}`;
  const response = await fetch(url);
  if (!response.ok && response.status !== 0) throw new Error('The selected Word document could not be opened.');
  const documentXml = unzipSync(new Uint8Array(await response.arrayBuffer()))['word/document.xml'];
  if (!documentXml) throw new Error('This is not a readable .docx Word document.');
  const paragraphs = (strFromU8(documentXml).match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) ?? []).map(paragraph =>
    (paragraph.replace(/<w:tab\s*\/?\s*>/g, '\t').replace(/<w:(?:br|cr)(?:\s[^>]*)?\/?\s*>/g, '\n').match(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>|\t|\n/g) ?? [])
      .map(run => run === '\t' || run === '\n' ? run : run.replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&')).join(''));
  const text = paragraphs.join('\n').trim();
  if (!text) throw new Error('No editable text was found in this Word document.');
  return {text, paragraphs: paragraphs.length};
};

export default function App() {
  const [message, setMessage] = useState('Choose a Word document to insert its text on the current note page.');
  const [busy, setBusy] = useState(false);
  const [position, setPosition] = useState({x: 0.5, y: 0.5});

  const importWordText = async () => {
    if (busy) return;
    setBusy(true); setMessage('Opening the Word document picker…');
    try {
      const selected: any = await RattaFileSelector.selectFile({selectType: 1, suffixList: ['docx'], maxNum: 1, title: 'Select Word document'});
      if (!selected || selected.length === 0) { setMessage('No document selected.'); return; }
      setMessage('Reading the Word document…');
      const result = await extractText(selected[0]);
      setMessage('Inserting editable text into the note…');
      const pathResponse: any = await PluginCommAPI.getCurrentFilePath();
      const pageResponse: any = await PluginCommAPI.getCurrentPageNum();
      const notePath = pathResponse?.success ? pathResponse.result : '';
      const page = pageResponse?.success ? pageResponse.result : 0;
      if (!notePath || !notePath.toLowerCase().endsWith('.note')) throw new Error('Open a Supernote note before importing the Word document.');
      const sizeResponse: any = await PluginFileAPI.getPageSize(notePath, page);
      const width = sizeResponse?.success ? sizeResponse.result.width : 1920;
      const height = sizeResponse?.success ? sizeResponse.result.height : 2560;
      const marginX = width * 0.04, marginY = height * 0.04, boxWidth = width * 0.72, boxHeight = height * 0.32;
      const left = Math.max(marginX, Math.min(width - marginX - boxWidth, position.x * width - boxWidth / 2));
      const top = Math.max(marginY, Math.min(height - marginY - boxHeight, position.y * height - boxHeight / 2));
      const insertion: any = await PluginNoteAPI.insertText({fontSize: 48, textContentFull: result.text, textRect: {left: Math.round(left), top: Math.round(top), right: Math.round(left + boxWidth), bottom: Math.round(top + boxHeight)}, textAlign: 0, textBold: 0, textItalics: 0, textFrameWidthType: 0, textEditable: 0});
      if (!insertion?.success) { Clipboard.setString(result.text); throw new Error('Supernote refused direct insertion. The extracted text is on the clipboard for pasting.'); }
      setMessage(`Success: ${result.paragraphs} paragraphs inserted as editable text.`);
    } catch (err: any) { setMessage('Error: ' + (err?.message ?? 'unknown error')); }
    finally { setBusy(false); }
  };

  return <View style={styles.container}>
    <Text style={styles.title}>Word Text Import</Text><Text style={styles.version}>Beta build 0.1.9</Text>
    <Text style={styles.positionLabel}>Tap the miniature page where the text box should go:</Text>
    <Pressable style={styles.pagePreview} onPress={({nativeEvent: {locationX, locationY}}) => setPosition({x: Math.max(0, Math.min(1, locationX / 180)), y: Math.max(0, Math.min(1, locationY / 240))})}>
      <View pointerEvents="none" style={[styles.boxPreview, {left: Math.max(7, Math.min(43, 180 * position.x - 65)), top: Math.max(8, Math.min(152, 240 * position.y - 40))}]}><Text style={styles.boxPreviewText}>TEXT</Text></View>
    </Pressable>
    <Text style={styles.message}>{message}</Text>
    <TouchableOpacity style={styles.button} onPress={importWordText} disabled={busy}><Text style={styles.buttonText}>{busy ? 'Working…' : 'Choose and insert Word document'}</Text></TouchableOpacity>
    <TouchableOpacity style={[styles.button, styles.closeButton]} onPress={() => PluginManager.closePluginView().catch(() => {})}><Text style={styles.buttonText}>Close</Text></TouchableOpacity>
  </View>;
}

const styles = StyleSheet.create({container: {flex: 1, padding: 24, justifyContent: 'center', backgroundColor: '#fff'}, title: {fontSize: 22, fontWeight: '600', marginBottom: 4, color: '#000', backgroundColor: '#fff'}, version: {fontSize: 13, marginBottom: 20, color: '#555', backgroundColor: '#fff'}, positionLabel: {fontSize: 14, marginBottom: 8, color: '#000', backgroundColor: '#fff'}, pagePreview: {width: 180, height: 240, borderWidth: 2, borderColor: '#000', backgroundColor: '#fff', alignSelf: 'center', marginBottom: 18}, boxPreview: {position: 'absolute', width: 130, height: 80, borderWidth: 2, borderColor: '#000', backgroundColor: '#ddd', alignItems: 'center', justifyContent: 'center'}, boxPreviewText: {fontSize: 13, fontWeight: '600', color: '#000'}, message: {fontSize: 15, marginBottom: 24, color: '#000', backgroundColor: '#fff'}, button: {backgroundColor: '#fff', borderWidth: 3, borderColor: '#000', paddingVertical: 14, paddingHorizontal: 10, alignItems: 'center', marginBottom: 12}, closeButton: {backgroundColor: '#fff', borderColor: '#067000'}, buttonText: {fontSize: 16, fontWeight: '600', color: '#000'}});
