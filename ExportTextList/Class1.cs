/*
 * Vegas Proで動作する自作スクリプト。プロジェクトのイベントトラックのうち、テキストを全部出力します。
 * Scripting code for Vegas Pro. Exporting All text on event tracks of tracks.
 * 
 * Created on 2020/10/15.
 * christinayan01 by Takahiro Yanai.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using ScriptPortal.Vegas;

namespace vegastest1 {
    public class EntryPoint {
        public void FromVegas(Vegas vegas) {

            // タイムライン上に何もないならやらない
            if (vegas.Project.MediaPool.Count == 0) {
                return;
            }

            // Common dialog.
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = "ExportTextList" + System.IO.Path.GetFileNameWithoutExtension(vegas.Project.FilePath) + ".txt";
            sfd.InitialDirectory = System.IO.Path.GetDirectoryName(vegas.Project.FilePath) + "\\";
            sfd.Filter = "テキストファイル(*.txt)|*.txt";
            if (sfd.ShowDialog() != DialogResult.OK) {
                return;
            }

            // Vegasプロジェクト設定の一時フォルダに一時ファイルを出力するためのファイル名
            string tempRtfFileName = vegas.Project.Video.PrerenderedFilesFolder + "temp.rtf";

            // タイムライン上にあるテキストのイベントを全部集める
            // ただしトラック毎にDictionaryで格納する
            Dictionary<long, Dictionary<long, string>> trackTable = new Dictionary<long, Dictionary<long, string>>();   // トラック毎の箱
            foreach (Track track in vegas.Project.Tracks) {
                Dictionary<long, string> table = new Dictionary<long, string>();    // 1つのトラックのテキストリスト。キーは開始位置。値はテキスト
                foreach (TrackEvent trackEvent in track.Events) {
                    foreach (Take take in trackEvent.Takes) {
                        string plainText = GetPlainText(take.Media, tempRtfFileName);
                        if (plainText.Length > 0) {
                            table.Add(trackEvent.Start.FrameCount, plainText);
                        }
                    }
                }

                trackTable.Add(track.Index, table);
           }

            // テキストファイルに出力
            // トラック毎にイベントの開始時間が若い順序で出力します
            System.IO.StreamWriter writer = new System.IO.StreamWriter(sfd.FileName, false, Encoding.GetEncoding("Shift_JIS"));
            foreach (var outputTrack in trackTable) {
                foreach (Track vegasTrack in vegas.Project.Tracks) {
                    if (vegasTrack.Index == outputTrack.Key) {
                        // 見出し部分を出力
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
                writer.WriteLine("");   // 1行空ける
            }
            writer.Close();

            MessageBox.Show("終了しました。");
        }

        // Vegas用テキストを得る
        string GetPlainText(Media media, string file) {
            string plainText = "";
            if (media.Generator != null) {
                // OFXEffect形式に変換して、Text情報を取得
                OFXEffect ofxEffect = media.Generator.OFXEffect;
                OFXStringParameter textParam = (OFXStringParameter)ofxEffect.FindParameterByName("Text");
                if (textParam != null) {
                    string rtfData = textParam.Value;   // これがテキスト

                    // まず、temp.rtfを書き出します
                    System.IO.StreamWriter writer = new System.IO.StreamWriter(file);
                    writer.WriteLine(rtfData);
                    writer.Close();

                    // 次に、temp.rtfを読み込むとプレーンテキストになっています
                    System.Windows.Forms.RichTextBox richtextBox = new System.Windows.Forms.RichTextBox();
                    richtextBox.Rtf = System.IO.File.ReadAllText(file);
                    plainText = richtextBox.Text;

                    // temp.rtfを削除
                    if (System.IO.File.Exists(file)) {
                        System.IO.File.Delete(file);
                    }
                }
            }
            return plainText;
        }
    }
}
