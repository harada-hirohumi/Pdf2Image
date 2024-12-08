using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using Windows.Data.Pdf;
using Windows.Storage.Streams;

namespace Pdf2Image
{
	/// <summary>
	/// PDFから画像変換
	/// </summary>
	internal class Program
	{
		/// <summary>
		/// プログラムのエントリポイント
		/// </summary>
		/// <param name="args">コマンドライン引数</param>
		static void Main(string[] args)
		{
			if (args.Length == 0)
			{
				Usage();
				return;
			}
			string filePath = args[0];
			ImageFormat format = DeterminFormat(args);
			if (!File.Exists(filePath) || Path.GetExtension(filePath).ToLower() != ".pdf" || null == format)
			{
				Usage();
				return;
			}
			Task task = ConvertAsync(args[0], format);
			task.Wait();
		}

		/// <summary>
		/// 使用方法
		/// </summary>
		private static void Usage()
		{
			Console.WriteLine("usage");
			Console.WriteLine("prompt>Pdf2Image InputFile FileType(optional)");
			Console.WriteLine("InputFile must be pdf");
			Console.WriteLine(string.Format("FileType:png or bmp or jpg or jpeg or gif or emf or wmf. default is {0}", Properties.Settings.Default.DefaultImageFomat));
		}

		/// <summary>
		/// 拡張子からImageFormatを確定する
		/// </summary>
		/// <param name="args">引数</param>
		/// <returns>ImageFormat</returns>
		private static ImageFormat DeterminFormat(string[] args)
		{
			string format;
			if (args.Length == 1)
			{
				format = Properties.Settings.Default.DefaultImageFomat.ToLower();
			}
			else
			{
				format = args[1].ToLower();
			}
			switch (format)
			{
				case "png":
					return ImageFormat.Png;
				case "bmp":
					return ImageFormat.Bmp;
				case "jpg":
				case "jpeg":
					return ImageFormat.Jpeg;
				case "gif":
					return ImageFormat.Gif;
				case "emf":
					return ImageFormat.Emf;
				case "wmf":
					return ImageFormat.Wmf;
				default:
					return null;
			}
		}

		/// <summary>
		/// 出力ファイル名のフォーマット
		/// </summary>
		/// <param name="nameBase">ファイル名の主部</param>
		/// <param name="index">インデックス</param>
		/// <param name="extension">拡張子</param>
		/// <returns>ファイル名</returns>
		private static string FormatFileName(string nameBase, uint index, string extension)
		{
			return string.Format("{0}_{1}.{2}", nameBase, index, extension);
		}

		/// <summary>
		/// 変換（非同期）
		/// 同期処理にしても良いけれど、複数ファイル取り扱いに改造する場合、非同期の方が良い
		/// </summary>
		/// <param name="filePath">入力ファイルパス</param>
		/// <param name="format">出力ファイルフォーマット</param>
		/// <returns>待機用Task</returns>
		private static async Task ConvertAsync(string filePath, ImageFormat format)
		{
			string nameBase = Path.GetFileNameWithoutExtension(filePath);
			string dirPath = Path.GetDirectoryName(filePath);
			string extension = format.ToString().ToLower();

			using (FileStream stream = File.OpenRead(filePath))
			using (IRandomAccessStream raStream = stream.AsRandomAccessStream())
			{
				PdfDocument document = await PdfDocument.LoadFromStreamAsync(raStream);

				for (uint index = 0; index < document.PageCount; index++)
				{
					string fileName = Path.Combine(dirPath, FormatFileName(nameBase, index, extension));

					// ページ出力を並行実行は、安全の為やめておく
					using (PdfPage page = document.GetPage(index))
					using (Stream outStream = new MemoryStream())
					using (IRandomAccessStream renderStream = outStream.AsRandomAccessStream())
					{
						// ページを出力ストリームに書き込み
						await page.RenderToStreamAsync(renderStream);

						// ビットマップに変換
						using (Bitmap bitmap = new Bitmap(outStream))
						{
							// 保存
							bitmap.Save(fileName, format);
						}
					}
				}
			}
		}
	}
}
