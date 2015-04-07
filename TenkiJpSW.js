/********************************************************************/
// TenkiJpSW.jp : 2005/01/28版 : kabao
// ◆DrT8mielhw氏のTenkiJp.js : 2004/02/23版をもとに作成しました。
//
/********************************************************************/

/********************************************************************/
// 変更必要設定ここから

// tenki.jpのURLの地域を表す数字部分
var cfgTenkiJpUrlBase = "45";

// ファイルの出力先。
// Schwatch.exeのあるフォルダを書いて下さい。
// パスの最後には \\ を付けてください。
// 注意：\ は \\ と書いてください。
// 例："C:\\Program Files\\Schwatch\\"
var cfgOutputDir = "";


/********************************************************************/
// 変更不要設定ここから

// true になっているファイルのみ作成
// 両方必ずtrueとする。
var cfgUseYoh		= true;		// 予報
var cfgUseYohWeek	= true;		// 週間予報

// 各データの前後に付ける文字。Prefixが左側、Suffixが右側に付く。
var cfgFmtKionHPrefix	= "";		// 最高気温
var cfgFmtKionHSuffix	= "℃";
var cfgFmtKionLPrefix	= "";		// 最低気温
var cfgFmtKionLSuffix	= "";
var cfgFmtKosuiKakurituPrefix	= "";	// 降水確率
var cfgFmtKosuiKakurituSuffix	= "％";

// 予報の "晴れ"、"くもり" などを置換える
// true で有効。
var cfgFmtYohRep = true;
var cfgFmtYohRepStr = new Array(
	"1",		// 晴れ
	"2",		// くもり
	"3",		// 雨
	"4",		// 雪
	"3",	// 雨か雪
	"4"		// 雪か雨
	);

// 予報の "のち"、"時々" などを置換える
var cfgFmtYohRep2 = true;
var cfgFmtYohRep2Str = new Array(
	"1",	// のち
	"1",	// のち時々
	"1",	// のち一時
	"2",	// 時々
	"2"	// 一時
	);
	

// 時刻の書式（書き方はReadme.txtを参照）
var cfgFmtYohDate1	= "%M%月%d%日(%dddj%)%hh%時発表";	// 予報の発表時刻
var cfgFmtYohDate2	= "%d%(%dddj%)";	// 各日付
var cfgReturnFmt	= "%hh%:%mm%";		// 関数の返却値


// ダウンロードオプション
var cfgDlNoCache	= false;	// true にするとキャッシュからの取得禁止

// メッセージボックスが自動的に閉じるまでの待ち時間を秒数で指定できます。
// 0 にすると自動的には閉じません。
// （単体実行時のメッセージボックスに対しても効きます）
var cfgPopupMsgSecondsToWait = 0;


// 単体実行時のメッセージモード
// 0 = エラーも含めて一切表示しない。
// 1 = エラーメッセージのみ表示する。
// 2 = 正常終了時も含めて必ず表示する。
var cfgStandAloneMsgMode = 1;

// ファイル名
var cfgFnTemp		= "_TenkiJpTemporary.txt";// 一時ファイル

// 設定ここまで
/********************************************************************/


/********************************************************************/
// エラーメッセージ。一応変更できます
var errDownload		= "ダウンロードエラー";
var errSaveToFile	= "に書き込めません";
var errOutFileCreate	= "出力先ファイルの作成に失敗";
var errTfssError	= "TFSSエラー:";


/********************************************************************/
// これより下はスクリプト本体
var fso		= new ActiveXObject("Scripting.FileSystemObject");
var WshShell	= new ActiveXObject("WScript.Shell");

var gPopupWndTitle = "TenkiJp.js";
var gStandAlone = false;

var gIsPopupErrMsg = false;


var gYohRep1 = new Array(
	/晴れ/,
	/くもり/,
	/雨/,
	/雪/,
	/雨か雪/,
	/雪か雨/
	);

var gYohRep2 = new Array(
	/\s+のち\s+/,
	/\s+のち時々\s*/,
	/\s+のち一時\s+/,
	/\s+時々\s+/,
	/\s+一時\s+/
	);

var cfgFnDay;
var cfgFnDayNext;

/********************************************************************/
function TenkiJp()
{
	var date = new Date();
	var strRet = date.FormatString(cfgReturnFmt);
	try
	{
		if(gStandAlone && cfgStandAloneMsgMode)
		{
			gIsPopupErrMsg = true;
		}

		// 出力先パスを設定
		if(!fso.FolderExists(cfgOutputDir))
		{
			if(WshShell.Popup(
				"ファイルは\n" + 
				cfgOutputDir + 
				"\nに作られます。\nこのフォルダを作成します。",
				0, gPopupWndTitle, 1 | 64) == 1)
			{
				fso.CreateFolder(cfgOutputDir);
			}
			else
				return;
		}

		// 各ファイルの出力先を設定
		var tempDate = new Date();
		var tempDateComp = new Date();
		tempDateComp.setTime(tempDate.getTime() + 604800000); //一週間後
		
		var tempMonth = "" + (tempDate.getMonth() + 1);
		if (tempMonth.length != 2) {
			tempMonth = "0" + tempMonth;
		}
		
		var tempMonthNext = "";
		// 必要なときのみ内容を入れる
		if (tempDateComp.getMonth() > tempDate.getMonth()) {
			tempMonthNext = tempMonthNext + (tempDate.getMonth() + 2);
			if (tempMonthNext.length != 2) {
				tempMonthNext = "0" + tempMonthNext;
			}
		}

		cfgFnDay = cfgOutputDir + tempDate.getYear() + tempMonth + ".day";
		if (tempMonthNext.length == 0) {
			cfgFnDayNext = "";
		} else {
			cfgFnDayNext = cfgOutputDir + tempDate.getYear() + tempMonthNext  + ".day";
		}
		cfgFnTemp = cfgOutputDir + cfgFnTemp;

		// URL
		var YohUrl = "http://tenki.jp/yoh/y" + cfgTenkiJpUrlBase + ".html";

		// tenki.jp からダウンロード＆出力
		// 予報
		if(cfgUseYoh || cfgUseYohWeek)
		{
			Download(YohUrl, cfgFnTemp, "");
			OutYohou(cfgFnTemp, cfgFnDay, cfgFnDayNext);
		}

	}
	catch(e)
	{
		if(gIsPopupErrMsg)
		{
			WshShell.Popup(e.description, cfgPopupMsgSecondsToWait,
				gPopupWndTitle, 16);
		}
		return e.description;
	}

	if(gStandAlone && (2 <= cfgStandAloneMsgMode))
		WshShell.Popup(strRet, cfgPopupMsgSecondsToWait, gPopupWndTitle, 0);
	return strRet;
}


/********************************************************************/
function Yohou()
{
	this.Date	= "";
	this.Youbi	= "";
	this.Tenki	= "";
	this.KionH	= "";
	this.KionL	= "";
	this.KosuiKakuritu = new Array();
	return this;
}


/********************************************************************/
// 予報を出力
function OutYohou(InputPath, OutPath, OutPathNext)
{
	try
	{
		var TenkiJpFile	= fso.OpenTextFile(InputPath, 1, false);

		// dayファイルを読み込みモードで開く。なければ作る。
		var outDateAndTenki = fso.OpenTextFile(OutPath, 1, true);
		var outDateAndTenkiNext = null;
		if (OutPathNext.length != 0) {
			outDateAndTenkiNext = fso.OpenTextFile(OutPathNext, 1, true);
		}
	}
	catch(e)
	{
		throw new Error(0, errOutFileCreate);
	}

	try
	{
		// 今日と明日の予報を読み取る
		var yohou = new Array(8);
		for(var n = 0; n < yohou.length; n++)
			yohou[n] = new Yohou();
		
		while(!TenkiJpFile.AtEndOfStream)
		{
			if("<!-- tenki1 -->" != TenkiJpFile.ReadLine())
				continue;
			break;
		}
		OutYohouHlp(TenkiJpFile, false, yohou);

		// 週間天気を読み取る
		if(cfgUseYohWeek)
		{
			while(!TenkiJpFile.AtEndOfStream)
			{
				if("<!-- tenki2 -->" != TenkiJpFile.ReadLine())
					continue;
				break;
			}
			OutYohouHlp(TenkiJpFile, true, yohou);
		}

		// 現在のdayファイルの内容を読み込み
		var strArr = new Array(32);
		for (var n = 0; n < 33; n++) {
			if (!outDateAndTenki.AtEndOfStream) {
				strArr[n] = outDateAndTenki.ReadLine();
			} else {
				strArr[n] = "00";
			}
		}
		outDateAndTenki.Close();
		
		// 翌月のdayファイルの内容を読み込み
		var strArrNext = null;
		if (outDateAndTenkiNext != null) {
			strArrNext = new Array(32);
			for (var n = 0; n < 33; n++) {
				if (!outDateAndTenkiNext.AtEndOfStream) {
					strArrNext[n] = outDateAndTenkiNext.ReadLine();
				} else {
					strArrNext[n] = "00";
				}
			}
			outDateAndTenkiNext.Close();
		}

		// 当月の予報をstrArrに書き込み
		var len = cfgUseYohWeek ? yohou.length : 2;
		for (var n = 0; n < len; n++)
		{
			if (yohou[n].Date != "" && parseInt(yohou[n].Date) >= parseInt(yohou[0].Date)) {
				strArr[yohou[n].Date] = strArr[yohou[n].Date].substring(0, 2) + yohou[n].Tenki;
			}
		}
		
		// 翌月の予報をstrArrNextに書き込み
		if (outDateAndTenkiNext != null) {
			len = cfgUseYohWeek ? yohou.length : 2;
			for (var n = 0; n < len; n++) {
				if (yohou[n].Date != "" && parseInt(yohou[n].Date) < parseInt(yohou[0].Date)) {
					strArrNext[yohou[n].Date] = strArrNext[yohou[n].Date].substring(0, 2) + yohou[n].Tenki;
				}
			}
		}
		
		// 当月出力
		outDateAndTenki = fso.OpenTextFile(OutPath, 2, true);
		for (var n = 0; n < 33; n++) {
			outDateAndTenki.WriteLine(strArr[n]);
		}
		
		// 翌月出力
		if (outDateAndTenkiNext != null) {
			outDateAndTenkiNext = fso.OpenTextFile(OutPathNext, 2, true);
			for (var n = 0; n < 33; n++) {
				outDateAndTenkiNext.WriteLine(strArrNext[n]);
			}
		}
		
	}
	catch(e)
	{
		throw e;
	}
	finally
	{
		TenkiJpFile.Close();
		outDateAndTenki.Close();
		if (outDateAndTenkiNext != null) {
			outDateAndTenkiNext.Close();
		}
	}
}

function OutYohouHlp(TenkiJpFile, Week, yohou)
{
	var LoopStart;
	var LoopEnd;
	if(Week)
	{
		LoopStart	= 2;
		LoopEnd		= 8;
	}
	else
	{
		LoopStart	= 0;
		LoopEnd		= 2;
	}
	
	var str = "";
	// 日付
	while(!TenkiJpFile.AtEndOfStream)
	{
		str = TenkiJpFile.ReadLine();
		if(-1 == str.indexOf("\"#333333\">日付</font>"))
			continue;

		for(var n = LoopStart; n < LoopEnd; n++)
		{
			str = TenkiJpFile.ReadLine();

			yohou[n].Date  = str.LTSlice("\">", "（");
			yohou[n].Date = yohou[n].Date.replace(/ /g,'');
			yohou[n].Youbi = str.LTSlice("（", "）</font>");

			var date = new Date();
			date.setDate(yohou[n].Date);
		}
		break;
	}

	// 天気
	while(!TenkiJpFile.AtEndOfStream)
	{
		str = TenkiJpFile.ReadLine();
		if(-1 == str.indexOf("\"#333333\">天気</font>"))
			continue;

		for(var n = LoopStart; n < LoopEnd; n++)
		{
			str = TenkiJpFile.ReadLine();
			yohou[n].Tenki  = str.LTSlice("alt=\"", "\"></td>");
			
			// 晴れ、くもり 等を置換え
			if(cfgFmtYohRep)
			{
				for(var i = gYohRep1.length-1; i >= 0; i--)
					yohou[n].Tenki = yohou[n].Tenki.replace(gYohRep1[i], cfgFmtYohRepStr[i]);
			}
			
			// のち、時々 等を置換え
			if(cfgFmtYohRep2)
			{
				for(var i = 0; i < gYohRep2.length; i++) {
					yohou[n].Tenki = yohou[n].Tenki.replace(gYohRep2[i], cfgFmtYohRep2Str[i]);
				}
			}
			
			// 5桁にする
			while (yohou[n].Tenki.length < 3) {
				yohou[n].Tenki = yohou[n].Tenki + "0";
			}
		}
		break;
	}
	
	// 気温
	while(!TenkiJpFile.AtEndOfStream)
	{
		str = TenkiJpFile.ReadLine();
		if(-1 == str.indexOf("#FF2200>最高</font>"))
			continue;

		for(var n = LoopStart; n < LoopEnd; n++)
		{
			str = TenkiJpFile.ReadLine();
			yohou[n].KionH  = str.LTSlice("#FF2200>", "</font>");
		}
		break;
	}
	while(!TenkiJpFile.AtEndOfStream)
	{
		str = TenkiJpFile.ReadLine();
		if(-1 == str.indexOf("\"#3333ff\">最低</font>"))
			continue;

		for(var n = LoopStart; n < LoopEnd; n++)
		{
			str = TenkiJpFile.ReadLine();
			yohou[n].KionL = str.LTSlice("\"#3333ff\">", "</font>");
		}
		break;
	}

	// 降水確率
	while(!TenkiJpFile.AtEndOfStream)
	{
		str = TenkiJpFile.ReadLine();
		if(-1 == str.indexOf("\"#333333\">降水"))
			continue;

		var k;
		if(Week)
			k = 1;
		else
			k = 4;

		var j = 0;
		var rowspan = new Array(0, 0);
		while((!TenkiJpFile.AtEndOfStream) && (j < k))
		{
			if(!Week)
			{
				str = TenkiJpFile.ReadLine();
				if(-1 == str.indexOf("\"#333333\">"))
					continue;
			}
			for(var n = LoopStart; n < LoopEnd; n++)
			{
				rowspan[n]--;
				if(0 < rowspan[n])
					continue;

				str = TenkiJpFile.ReadLine();
				if(-1 != str.indexOf("rowspan=\""))
				{
					rowspan[n] = parseInt(str.TSlice("rowspan=\"", "\">"), 10);
					for(var i = j; i < (rowspan[n]+j); i++)
					{
						yohou[n].KosuiKakuritu[i] = str.LTSlice("\"#FF6600\">", "</font>");
					}
				}
				else
				{
					yohou[n].KosuiKakuritu[j] = str.LTSlice("\"#FF6600\">", "</font>");
				}
			}
			j++;
		}
		break;
	}

	// 気温、降水確率の余計な文字を削除し、
	// Prefix、Suffixを付ける。
	for(var n = LoopStart; n < LoopEnd; n++)
	{
		yohou[n].KionH = yohou[n].KionH.LTSliceLeft("℃");
		if("-" != yohou[n].KionH)
		{
			yohou[n].KionH = cfgFmtKionHPrefix + yohou[n].KionH + cfgFmtKionHSuffix;
		}

		yohou[n].KionL = yohou[n].KionL.LTSliceLeft("℃");
		if("-" != yohou[n].KionL)
		{
			yohou[n].KionL= cfgFmtKionLPrefix + yohou[n].KionL + cfgFmtKionLSuffix;
		}

		for(var j = 0; j < yohou[n].KosuiKakuritu.length; j++)
		{
			yohou[n].KosuiKakuritu[j] = yohou[n].KosuiKakuritu[j].LTSliceLeft("％");
			if("-" != yohou[n].KosuiKakuritu[j])
			{
				yohou[n].KosuiKakuritu[j] =
					cfgFmtKosuiKakurituPrefix + yohou[n].KosuiKakuritu[j] + cfgFmtKosuiKakurituSuffix;
			}
		}
	}
}

/********************************************************************/
// Url をダウンロードし、OutputFilePath に書き込む。
// LastMod を "" 以外にすると、
// Last-Modified が一致する場合ダウンロードしない。
// Last-Modified を返す。
function Download(Url, OutputFilePath, LastMod)
{
//	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var xmlhttp;
	var ResponseLastMod = "";
	try
	{
		xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
		if("" != LastMod)
		{
			xmlhttp.open("HEAD", Url, false);
			if(cfgDlNoCache)
			{
				xmlhttp.setRequestHeader("Pragma", "no-cache");
				xmlhttp.setRequestHeader("Cache-Control", "no-cache");
			}
			xmlhttp.send();

			ResponseLastMod = xmlhttp.getResponseHeader("Last-Modified");
			if(LastMod == ResponseLastMod)
				return ResponseLastMod;
		}

		xmlhttp.open("GET", Url, false);
		if(cfgDlNoCache)
		{
			xmlhttp.setRequestHeader("Pragma", "no-cache");
			xmlhttp.setRequestHeader("Cache-Control", "no-cache");
		}
		xmlhttp.send();
		ResponseLastMod = xmlhttp.getResponseHeader("Last-Modified");
	}
	catch(e)
	{
		throw new Error(0, errDownload);
	}

	if(200 != xmlhttp.status)
	{
		throw new Error(0, xmlhttp.statusText);
	}

	try
	{
		var stream = new ActiveXObject("ADODB.Stream");
		stream.Mode = 3;
		stream.Type = 1;
		stream.Open();
		stream.Write(xmlhttp.responseBody);
		stream.SaveToFile(OutputFilePath, 2);
		stream.Close();
	}
	catch(e)
	{
		throw new Error(0, OutputFilePath + errSaveToFile);
	}

	EucToShiftJis(OutputFilePath, OutputFilePath);
	return ResponseLastMod;
}


/********************************************************************/
function EucToShiftJis(input, output)
{
	// EUC -> SJISへ変換
	var adodbstreamLoad = new ActiveXObject("ADODB.Stream");
	adodbstreamLoad.Open();
	adodbstreamLoad.Type = 2;
	adodbstreamLoad.Charset = "EUC-JP";
	var adodbstreamSave = new ActiveXObject("ADODB.Stream");
	adodbstreamSave.Open();
	adodbstreamSave.Type = 2;
	adodbstreamSave.Charset = "SHIFT-JIS";

	adodbstreamLoad.LoadFromFile(input);
	adodbstreamLoad.CopyTo(adodbstreamSave);
	adodbstreamLoad.Close();
	adodbstreamSave.SaveToFile(output, 2);
	adodbstreamSave.Close();
}


/********************************************************************/
// 文字列 p1 と p2 の間の文字列を返す。左から右へ検索
String.prototype.TSlice = function(p1, p2)
{
	var pos = this.indexOf(p1) + p1.length;
	if(-1 == pos)
		return "";
	return this.slice(pos, this.indexOf(p2, pos));
}
// 文字列 p1 と p2 の間の文字列を返す。右から左へ検索
String.prototype.LTSlice = function(p1, p2)
{
	var pos = this.lastIndexOf(p2);
	if(-1 == pos)
		return "";
	return this.slice(this.lastIndexOf(p1, pos) + p1.length, pos);
}


// 文字列 p1 を除いた、左側を返す。右から検索
String.prototype.LTSliceLeft = function(p1)
{
	var pos = this.lastIndexOf(p1);
	if(-1 == pos)
		return this;
	return this.slice(0, pos);
}


// 文字列 p1 を除いた、右側を返す。左から検索
String.prototype.TSliceRight = function(p1)
{
	var pos = this.indexOf(p1) + p1.length;
	if(-1 == pos)
		return this;
	return this.slice(pos);
}



/********************************************************************/
var gMonthArrayE = new Array(
  new Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"),
  new Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
);
var gMonthArrayJ = new Array(
  new Array("睦月", "如月", "弥生", "卯月", "皐月", "水無月", "文月", "葉月", "長月", "神無月", "霜月", "師走")
);
var gDayArrayE = new Array(
  new Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"),
  new Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
);
var gDayArrayJ = new Array(
  new Array("日", "月", "火", "水", "木", "金", "土"),
  new Array("日曜", "月曜", "火曜", "水曜", "木曜", "金曜", "土曜"),
  new Array("日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日")
);
Date.prototype.FormatString = function(format)
{
	var s = format;
	var a = s.match(/%[^%]*%/g);
	var v;
	var kan = 0;
	for(var n = 0; n < a.length; n++)
	{
		switch(a[n])
		{
		// ローカル
		// 年
		case "%yyyy%":	v = this.getFullYear();
			break;
		case "%yy%":	v = this.getFullYear().toString().match(/\d{2}$/);
			break;

		// 月
		case "%M%":	v = this.getMonth() + 1;
			break;
		case "%MM%":	v = ("0" + (this.getMonth() + 1)).match(/\d{2}$/);
			break;
		case "%MMME%":	v = gMonthArrayE[0][this.getMonth()];
			break;
		case "%MMMME%":	v = gMonthArrayE[1][this.getMonth()];
			break;
		case "%MMMJ%":	v = gMonthArrayJ[0][this.getMonth()];
			break;

		// 日
		case "%d%":	v = this.getDate();
			break;
		case "%dd%":	v = ("0" + this.getDate()).match(/\d{2}$/);
			break;
		case "%ddde%":	v = gDayArrayE[0][this.getDay()];
			break;
		case "%dddde%":	v = gDayArrayE[1][this.getDay()];
			break;
		case "%dddj%":	v = gDayArrayJ[0][this.getDay()];
			break;
		case "%ddddj%":	v = gDayArrayJ[1][this.getDay()];
			break;
		case "%dddddj%":v = gDayArrayJ[2][this.getDay()];
			break;

		// 時 24時間
		case "%h%":	v = this.getHours();
			break;
		case "%hh%":	v = ("0" + this.getHours()).match(/\d{2}$/);
			break;
		// 時 12時間
		case "%h12%":
			var x = this.getHours();
			v = x >= 12 ? x - 12 : x;
			break;
		case "%hh12%":
			var x = this.getHours();
			v = x >= 12 ? ("0"+(x - 12)).match(/\d{2}$/) : ("0" + x).match(/\d{2}$/);
			break;

		case "%ampm%":	v = this.getHours() >= 12 ? "pm" : "am";
			break;
		case "%AMPM%":	v = this.getHours() >= 12 ? "PM" : "AM";
			break;
		case "%ampmj%":	v = this.getHours() >= 12 ? "午後" : "午前";
			break;

		// 分
		case "%m%":	v = this.getMinutes();
			break;
		case "%mm%":	v = ("0" + this.getMinutes()).match(/\d{2}$/);
			break;

		// 秒
		case "%s%":	v = this.getSeconds();
			break;
		case "%ss%":	v = ("0" + this.getSeconds()).match(/\d{2}$/);
			break;

		// ミリ秒
		case "%z%":	v = this.getMilliseconds();
			break;
		case "%zzz%":	v =("000" + this.getMilliseconds()).match(/\d{3}$/);
			break;


		// インターネットタイム
		case "%inettime%":
			v = Math.round((((this.getUTCHours()+1) * 60 + this.getUTCMinutes())/1.44));
			break;

		// その他
		case "%_n%":	v = "\n";
			break;
		case "%%":	v = "%";
			break;
		
		default:
			continue;
		}

		s = s.replace(new RegExp(a[n]), v);
	}
	return s;
}


/********************************************************************/
// 単体実行
if("object" == typeof WScript)
{
	gStandAlone = true;
	TenkiJp();
}
