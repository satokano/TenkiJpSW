/********************************************************************/
// TenkiJpSW.jp : 2005/01/28�� : kabao
// ��DrT8mielhw����TenkiJp.js : 2004/02/23�ł����Ƃɍ쐬���܂����B
//
/********************************************************************/

/********************************************************************/
// �ύX�K�v�ݒ肱������

// tenki.jp��URL�̒n���\����������
var cfgTenkiJpUrlBase = "45";

// �t�@�C���̏o�͐�B
// Schwatch.exe�̂���t�H���_�������ĉ������B
// �p�X�̍Ō�ɂ� \\ ��t���Ă��������B
// ���ӁF\ �� \\ �Ə����Ă��������B
// ��F"C:\\Program Files\\Schwatch\\"
var cfgOutputDir = "";


/********************************************************************/
// �ύX�s�v�ݒ肱������

// true �ɂȂ��Ă���t�@�C���̂ݍ쐬
// �����K��true�Ƃ���B
var cfgUseYoh		= true;		// �\��
var cfgUseYohWeek	= true;		// �T�ԗ\��

// �e�f�[�^�̑O��ɕt���镶���BPrefix�������ASuffix���E���ɕt���B
var cfgFmtKionHPrefix	= "";		// �ō��C��
var cfgFmtKionHSuffix	= "��";
var cfgFmtKionLPrefix	= "";		// �Œ�C��
var cfgFmtKionLSuffix	= "";
var cfgFmtKosuiKakurituPrefix	= "";	// �~���m��
var cfgFmtKosuiKakurituSuffix	= "��";

// �\��� "����"�A"������" �Ȃǂ�u������
// true �ŗL���B
var cfgFmtYohRep = true;
var cfgFmtYohRepStr = new Array(
	"1",		// ����
	"2",		// ������
	"3",		// �J
	"4",		// ��
	"3",	// �J����
	"4"		// �Ⴉ�J
	);

// �\��� "�̂�"�A"���X" �Ȃǂ�u������
var cfgFmtYohRep2 = true;
var cfgFmtYohRep2Str = new Array(
	"1",	// �̂�
	"1",	// �̂����X
	"1",	// �̂��ꎞ
	"2",	// ���X
	"2"	// �ꎞ
	);
	

// �����̏����i��������Readme.txt���Q�Ɓj
var cfgFmtYohDate1	= "%M%��%d%��(%dddj%)%hh%�����\";	// �\��̔��\����
var cfgFmtYohDate2	= "%d%(%dddj%)";	// �e���t
var cfgReturnFmt	= "%hh%:%mm%";		// �֐��̕ԋp�l


// �_�E�����[�h�I�v�V����
var cfgDlNoCache	= false;	// true �ɂ���ƃL���b�V������̎擾�֎~

// ���b�Z�[�W�{�b�N�X�������I�ɕ���܂ł̑҂����Ԃ�b���Ŏw��ł��܂��B
// 0 �ɂ���Ǝ����I�ɂ͕��܂���B
// �i�P�̎��s���̃��b�Z�[�W�{�b�N�X�ɑ΂��Ă������܂��j
var cfgPopupMsgSecondsToWait = 0;


// �P�̎��s���̃��b�Z�[�W���[�h
// 0 = �G���[���܂߂Ĉ�ؕ\�����Ȃ��B
// 1 = �G���[���b�Z�[�W�̂ݕ\������B
// 2 = ����I�������܂߂ĕK���\������B
var cfgStandAloneMsgMode = 1;

// �t�@�C����
var cfgFnTemp		= "_TenkiJpTemporary.txt";// �ꎞ�t�@�C��

// �ݒ肱���܂�
/********************************************************************/


/********************************************************************/
// �G���[���b�Z�[�W�B�ꉞ�ύX�ł��܂�
var errDownload		= "�_�E�����[�h�G���[";
var errSaveToFile	= "�ɏ������߂܂���";
var errOutFileCreate	= "�o�͐�t�@�C���̍쐬�Ɏ��s";
var errTfssError	= "TFSS�G���[:";


/********************************************************************/
// �����艺�̓X�N���v�g�{��
var fso		= new ActiveXObject("Scripting.FileSystemObject");
var WshShell	= new ActiveXObject("WScript.Shell");

var gPopupWndTitle = "TenkiJp.js";
var gStandAlone = false;

var gIsPopupErrMsg = false;


var gYohRep1 = new Array(
	/����/,
	/������/,
	/�J/,
	/��/,
	/�J����/,
	/�Ⴉ�J/
	);

var gYohRep2 = new Array(
	/\s+�̂�\s+/,
	/\s+�̂����X\s*/,
	/\s+�̂��ꎞ\s+/,
	/\s+���X\s+/,
	/\s+�ꎞ\s+/
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

		// �o�͐�p�X��ݒ�
		if(!fso.FolderExists(cfgOutputDir))
		{
			if(WshShell.Popup(
				"�t�@�C����\n" + 
				cfgOutputDir + 
				"\n�ɍ���܂��B\n���̃t�H���_���쐬���܂��B",
				0, gPopupWndTitle, 1 | 64) == 1)
			{
				fso.CreateFolder(cfgOutputDir);
			}
			else
				return;
		}

		// �e�t�@�C���̏o�͐��ݒ�
		var tempDate = new Date();
		var tempDateComp = new Date();
		tempDateComp.setTime(tempDate.getTime() + 604800000); //��T�Ԍ�
		
		var tempMonth = "" + (tempDate.getMonth() + 1);
		if (tempMonth.length != 2) {
			tempMonth = "0" + tempMonth;
		}
		
		var tempMonthNext = "";
		// �K�v�ȂƂ��̂ݓ��e������
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

		// tenki.jp ����_�E�����[�h���o��
		// �\��
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
// �\����o��
function OutYohou(InputPath, OutPath, OutPathNext)
{
	try
	{
		var TenkiJpFile	= fso.OpenTextFile(InputPath, 1, false);

		// day�t�@�C����ǂݍ��݃��[�h�ŊJ���B�Ȃ���΍��B
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
		// �����Ɩ����̗\���ǂݎ��
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

		// �T�ԓV�C��ǂݎ��
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

		// ���݂�day�t�@�C���̓��e��ǂݍ���
		var strArr = new Array(32);
		for (var n = 0; n < 33; n++) {
			if (!outDateAndTenki.AtEndOfStream) {
				strArr[n] = outDateAndTenki.ReadLine();
			} else {
				strArr[n] = "00";
			}
		}
		outDateAndTenki.Close();
		
		// ������day�t�@�C���̓��e��ǂݍ���
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

		// �����̗\���strArr�ɏ�������
		var len = cfgUseYohWeek ? yohou.length : 2;
		for (var n = 0; n < len; n++)
		{
			if (yohou[n].Date != "" && parseInt(yohou[n].Date) >= parseInt(yohou[0].Date)) {
				strArr[yohou[n].Date] = strArr[yohou[n].Date].substring(0, 2) + yohou[n].Tenki;
			}
		}
		
		// �����̗\���strArrNext�ɏ�������
		if (outDateAndTenkiNext != null) {
			len = cfgUseYohWeek ? yohou.length : 2;
			for (var n = 0; n < len; n++) {
				if (yohou[n].Date != "" && parseInt(yohou[n].Date) < parseInt(yohou[0].Date)) {
					strArrNext[yohou[n].Date] = strArrNext[yohou[n].Date].substring(0, 2) + yohou[n].Tenki;
				}
			}
		}
		
		// �����o��
		outDateAndTenki = fso.OpenTextFile(OutPath, 2, true);
		for (var n = 0; n < 33; n++) {
			outDateAndTenki.WriteLine(strArr[n]);
		}
		
		// �����o��
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
	// ���t
	while(!TenkiJpFile.AtEndOfStream)
	{
		str = TenkiJpFile.ReadLine();
		if(-1 == str.indexOf("\"#333333\">���t</font>"))
			continue;

		for(var n = LoopStart; n < LoopEnd; n++)
		{
			str = TenkiJpFile.ReadLine();

			yohou[n].Date  = str.LTSlice("\">", "�i");
			yohou[n].Date = yohou[n].Date.replace(/ /g,'');
			yohou[n].Youbi = str.LTSlice("�i", "�j</font>");

			var date = new Date();
			date.setDate(yohou[n].Date);
		}
		break;
	}

	// �V�C
	while(!TenkiJpFile.AtEndOfStream)
	{
		str = TenkiJpFile.ReadLine();
		if(-1 == str.indexOf("\"#333333\">�V�C</font>"))
			continue;

		for(var n = LoopStart; n < LoopEnd; n++)
		{
			str = TenkiJpFile.ReadLine();
			yohou[n].Tenki  = str.LTSlice("alt=\"", "\"></td>");
			
			// ����A������ ����u����
			if(cfgFmtYohRep)
			{
				for(var i = gYohRep1.length-1; i >= 0; i--)
					yohou[n].Tenki = yohou[n].Tenki.replace(gYohRep1[i], cfgFmtYohRepStr[i]);
			}
			
			// �̂��A���X ����u����
			if(cfgFmtYohRep2)
			{
				for(var i = 0; i < gYohRep2.length; i++) {
					yohou[n].Tenki = yohou[n].Tenki.replace(gYohRep2[i], cfgFmtYohRep2Str[i]);
				}
			}
			
			// 5���ɂ���
			while (yohou[n].Tenki.length < 3) {
				yohou[n].Tenki = yohou[n].Tenki + "0";
			}
		}
		break;
	}
	
	// �C��
	while(!TenkiJpFile.AtEndOfStream)
	{
		str = TenkiJpFile.ReadLine();
		if(-1 == str.indexOf("#FF2200>�ō�</font>"))
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
		if(-1 == str.indexOf("\"#3333ff\">�Œ�</font>"))
			continue;

		for(var n = LoopStart; n < LoopEnd; n++)
		{
			str = TenkiJpFile.ReadLine();
			yohou[n].KionL = str.LTSlice("\"#3333ff\">", "</font>");
		}
		break;
	}

	// �~���m��
	while(!TenkiJpFile.AtEndOfStream)
	{
		str = TenkiJpFile.ReadLine();
		if(-1 == str.indexOf("\"#333333\">�~��"))
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

	// �C���A�~���m���̗]�v�ȕ������폜���A
	// Prefix�ASuffix��t����B
	for(var n = LoopStart; n < LoopEnd; n++)
	{
		yohou[n].KionH = yohou[n].KionH.LTSliceLeft("��");
		if("-" != yohou[n].KionH)
		{
			yohou[n].KionH = cfgFmtKionHPrefix + yohou[n].KionH + cfgFmtKionHSuffix;
		}

		yohou[n].KionL = yohou[n].KionL.LTSliceLeft("��");
		if("-" != yohou[n].KionL)
		{
			yohou[n].KionL= cfgFmtKionLPrefix + yohou[n].KionL + cfgFmtKionLSuffix;
		}

		for(var j = 0; j < yohou[n].KosuiKakuritu.length; j++)
		{
			yohou[n].KosuiKakuritu[j] = yohou[n].KosuiKakuritu[j].LTSliceLeft("��");
			if("-" != yohou[n].KosuiKakuritu[j])
			{
				yohou[n].KosuiKakuritu[j] =
					cfgFmtKosuiKakurituPrefix + yohou[n].KosuiKakuritu[j] + cfgFmtKosuiKakurituSuffix;
			}
		}
	}
}

/********************************************************************/
// Url ���_�E�����[�h���AOutputFilePath �ɏ������ށB
// LastMod �� "" �ȊO�ɂ���ƁA
// Last-Modified ����v����ꍇ�_�E�����[�h���Ȃ��B
// Last-Modified ��Ԃ��B
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
	// EUC -> SJIS�֕ϊ�
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
// ������ p1 �� p2 �̊Ԃ̕������Ԃ��B������E�֌���
String.prototype.TSlice = function(p1, p2)
{
	var pos = this.indexOf(p1) + p1.length;
	if(-1 == pos)
		return "";
	return this.slice(pos, this.indexOf(p2, pos));
}
// ������ p1 �� p2 �̊Ԃ̕������Ԃ��B�E���獶�֌���
String.prototype.LTSlice = function(p1, p2)
{
	var pos = this.lastIndexOf(p2);
	if(-1 == pos)
		return "";
	return this.slice(this.lastIndexOf(p1, pos) + p1.length, pos);
}


// ������ p1 ���������A������Ԃ��B�E���猟��
String.prototype.LTSliceLeft = function(p1)
{
	var pos = this.lastIndexOf(p1);
	if(-1 == pos)
		return this;
	return this.slice(0, pos);
}


// ������ p1 ���������A�E����Ԃ��B�����猟��
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
  new Array("�r��", "�@��", "�퐶", "�K��", "�H��", "������", "����", "�t��", "����", "�_����", "����", "�t��")
);
var gDayArrayE = new Array(
  new Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"),
  new Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
);
var gDayArrayJ = new Array(
  new Array("��", "��", "��", "��", "��", "��", "�y"),
  new Array("���j", "���j", "�Ηj", "���j", "�ؗj", "���j", "�y�j"),
  new Array("���j��", "���j��", "�Ηj��", "���j��", "�ؗj��", "���j��", "�y�j��")
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
		// ���[�J��
		// �N
		case "%yyyy%":	v = this.getFullYear();
			break;
		case "%yy%":	v = this.getFullYear().toString().match(/\d{2}$/);
			break;

		// ��
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

		// ��
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

		// �� 24����
		case "%h%":	v = this.getHours();
			break;
		case "%hh%":	v = ("0" + this.getHours()).match(/\d{2}$/);
			break;
		// �� 12����
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
		case "%ampmj%":	v = this.getHours() >= 12 ? "�ߌ�" : "�ߑO";
			break;

		// ��
		case "%m%":	v = this.getMinutes();
			break;
		case "%mm%":	v = ("0" + this.getMinutes()).match(/\d{2}$/);
			break;

		// �b
		case "%s%":	v = this.getSeconds();
			break;
		case "%ss%":	v = ("0" + this.getSeconds()).match(/\d{2}$/);
			break;

		// �~���b
		case "%z%":	v = this.getMilliseconds();
			break;
		case "%zzz%":	v =("000" + this.getMilliseconds()).match(/\d{3}$/);
			break;


		// �C���^�[�l�b�g�^�C��
		case "%inettime%":
			v = Math.round((((this.getUTCHours()+1) * 60 + this.getUTCMinutes())/1.44));
			break;

		// ���̑�
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
// �P�̎��s
if("object" == typeof WScript)
{
	gStandAlone = true;
	TenkiJp();
}
