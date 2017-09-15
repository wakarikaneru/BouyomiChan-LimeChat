////////////////////////////////////////////////////////////////////////////////////////////////////
// BouyomiLimeChat.js ～ 棒読みちゃん・LimeChat連携スクリプト
////////////////////////////////////////////////////////////////////////////////////////////////////
//■導入方法
// 1.当ファイルをLimeChatのscriptsフォルダに配置する
//   例）C:\【LimeChatインストール先】\users\【アカウント名】\scripts
//
// 2.LimeChat側でスクリプトを有効にする
//   ・LimeChatのメニューから「設定→スクリプトの設定」を開く。
//   ・スクリプトの設定画面で、「BouyomiLimeChat.js」の行を右クリックし、○を付ける。
//   ・スクリプトの設定画面の閉じるボタンを押す。
//
////////////////////////////////////////////////////////////////////////////////////////////////////
//■設定

//発言者の名前を読み上げるかどうか(true:読む, false:読まない)
var bNick = false;

//入出情報を読み上げるかどうか(true:読む, false:読まない)
var bInOut = false;


//声質の設定
//半角スペースで区切って入力してください([速度] [音程] [音量]) 
//   [速度]       -1:デフォルト、50:最低、300:最高
//   [音程]       -1:デフォルト、50:最低、200:最高
//   [音量]       -1:デフォルト、 0:最低、100:最高
var defaultVoiceParam = "-1 -1 -1";

//ランダムで声質を変更するか(true:変更する, false:変更しない)
var bRandomVoice = true;

////////////////////////////////////////////////////////////////////////////////////////////////////

var sRemoteTalkCmd = null;
var oShell;
var oWmi;

function addTalkTask(text, voice) {
	if(sRemoteTalkCmd == null) {
		findRemoteTalk();
		if(sRemoteTalkCmd == null) {
			log("RemoteTalkが見つからないのでスキップ-" + text);
			return;
		}
	}
	
	oShell.Run(sRemoteTalkCmd + " \"" + text.replace("\"", " ") + "\" " + voice, 0, false);
}

function talkChat(prefix, text) {
	var voiceParam;
	
	//ニックネームからランダムで声質を変更する
	if (bRandomVoice){
		//   [声質]        0:デフォルト、 1:女性1, 2:女性2, 3:男性1, 4:男性2,
		//                 5:中性, 6:ロボット, 7:機械1, 8:機械2, 9以降:SAPI
		var nick=prefix.nick;
		
		var hash = 0;
		for(var i = 0; i < nick.length; i++){
			hash += nick.charCodeAt(i);
		}
		hash = hash % 9;
		
		voiceParam = defaultVoiceParam + " " + hash;
	}else{
		voiceParam = defaultVoiceParam + " 0";
	}
	
	if (bNick){
		addTalkTask(prefix.nick + "。" + text, voiceParam);
	} else {
		addTalkTask(text, voiceParam);
	}
}

function findRemoteTalk() {
	var proc = oWmi.ExecQuery("Select * from Win32_Process Where Name like 'BouyomiChan.exe'");
	var e    = new Enumerator(proc);
	for(; !e.atEnd(); e.moveNext()) {
		var item = e.item();
		
		var path = item.ExecutablePath.replace("\\BouyomiChan.exe", "");
		sRemoteTalkCmd = "\"" + path + "\\RemoteTalk\\RemoteTalk.exe\" /T";
		
		log("棒読みちゃん検出:" + path);
	}
}

////////////////////////////////////////////////////////////////////////////////////////////////////

function event::onLoad() {
	oShell = new ActiveXObject("Wscript.Shell");
	oWmi   = GetObject("winmgmts:\\\\.\\root\\cimv2");
	
	//addTalkTask("ライムチャットとの連携を開始しました");
}

function event::onUnLoad() {
	oShell = null;
	oWmi   = null;
	
	//addTalkTask("ライムチャットとの連携を終了しました");
}

function event::onConnect(){
	addTalkTask(name + "サーバに接続しました");
}

function event::onDisconnect(){
	addTalkTask(name + "サーバから切断しました");
}

function event::onJoin(prefix, channel) {
	if (bInOut) {
		addTalkTask(prefix.nick + "さんが " + channel + " に入りました");
	}
}

function event::onPart(prefix, channel, comment) {
	if (bInOut) {
		addTalkTask(prefix.nick + "さんが " + channel + " から出ました。");
	}
}

function event::onQuit(prefix, comment) {
	if (bInOut) {
		addTalkTask(prefix.nick + "さんがサーバから切断しました。");
	}
}

function event::onChannelText(prefix, channel, text) {
	talkChat(prefix, text);
	//log("CnannelText[" + channel + "]" + text);
}

function event::onChannelNotice(prefix, channel, text) {
	talkChat(prefix, text);
	//log("CnannelNotice[" + channel + "]" + text);
}

function event::onChannelAction(prefix, channel, text) {
	talkChat(prefix, text);
	//log("CnannelAction[" + channel + "]" + text);
}

function event::onTalkText(prefix, targetNick, text) {
	talkChat(prefix, text);
	//log("TalkText[" + prefix.nick + "]" + text);
}

function event::onTalkNotice(prefix, targetNick, text) {
	talkChat(prefix, text);
	//log("TalkNotice[" + prefix.nick + "]" + text);
}

function event::onTalkAction(prefix, targetNick, text) {
	talkChat(prefix, text);
	//log("TalkAction[" + prefix.nick + "]" + text);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
