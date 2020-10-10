/*
 * Vegas Proで動作する自作スクリプト。プロジェクトのイベントトラックのうち、テキストを全部出力します。
 * Scripting code for Vegas Pro. Exporting All text on event tracks of tracks.
 * 
 * Created on 2020/10/10.
 * christinayan01 by Takahiro Yanai.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using ScriptPortal.Vegas;

namespace vegastest1 {
    public class EntryPoint {
        public void FromVegas(Vegas vegas) {

            int count = vegas.Project.MediaPool.Count;
            if (count == 0) {
                return;
            }

            string dir = vegas.Project.FilePath;
            dir = System.IO.Path.GetDirectoryName(dir) + "\\";

            // コモンダイアログ
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = "ExportTextList" + System.IO.Path.GetFileNameWithoutExtension(vegas.Project.FilePath) + ".txt";
            sfd.InitialDirectory = dir;
            sfd.Filter = "テキストファイル(*.txt)|*.txt";
            if (sfd.ShowDialog() != DialogResult.OK) {
                return;
            }

            // 一時ファイル名
            string tempRtfName = vegas.Project.Video.PrerenderedFilesFolder + "temp.rtf";

            // トラック毎の箱
            Dictionary<long, Dictionary<long, string>> trackTable = new Dictionary<long, Dictionary<long, string>>();

            foreach (Track track in vegas.Project.Tracks) {
                // トラック内のテキストリスト。キーは開始位置。値はテキスト
                Dictionary<long, string> table = new Dictionary<long, string>();

                foreach (TrackEvent trackEvent in track.Events) {
                    foreach (Take take in trackEvent.Takes) {
                        Media media = take.Media;
                        string originalStr = GetPlainText(media, tempRtfName);
                        if (originalStr.Length > 0) {
                            table.Add(trackEvent.Start.FrameCount, originalStr);
                        }
                    }
                }

                trackTable.Add(track.Index, table);
           }

            // 出力
            System.IO.StreamWriter writer = new System.IO.StreamWriter(sfd.FileName, false, Encoding.GetEncoding("Shift_JIS"));
            foreach (var outputTrack in trackTable) {

                foreach (Track vegasTrack in vegas.Project.Tracks) {
                    if (vegasTrack.Index == outputTrack.Key) {

                        // トラック番号
                        string trackName = "-";
                        if (vegasTrack.Name != null) {
                            trackName = vegasTrack.Name;
                        }
                        writer.WriteLine("[" + vegasTrack.DisplayIndex + ":" + trackName + "]");

                        // このトラックをタイムライン順に出力
                        Dictionary<long, string> table = outputTrack.Value;
                        foreach (var va in table) {

                            writer.WriteLine(va.Value);
                        }
                    }
                }
                writer.WriteLine("");
            }
            writer.Close();
        }

        // Vegas用テキストを得る
        string GetPlainText(Media media, string file) {

            string plainText = "";
            if (media.Generator != null) {
                // OFXEffect形式に変換して、Text情報を取得。
                OFXEffect ofxEffect = media.Generator.OFXEffect;
                OFXStringParameter textParam = (OFXStringParameter)ofxEffect.FindParameterByName("Text");
                if (textParam != null) {
                    string rtfData = textParam.Value;   // これがテキスト

                    // まず、RTFファイルとして書き出します
                    System.IO.StreamWriter writer = new System.IO.StreamWriter(file);
                    writer.WriteLine(rtfData);
                    writer.Close();

                    // 次に、そのRTFファイルを読み込んで、プレーンテキストに変換します
                    System.Windows.Forms.RichTextBox rtBox = new System.Windows.Forms.RichTextBox();
                    string s = System.IO.File.ReadAllText(file);
                    rtBox.Rtf = s;
                    plainText = rtBox.Text;

                    if (System.IO.File.Exists(file)) {
                        System.IO.File.Delete(file);
                    }
                }
            }
            return plainText;
        }
    }
}

