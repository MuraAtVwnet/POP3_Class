##############################################
# TCP Client Class
##############################################
class TCPClient {
	#-------------------------------------------------------------------------
	# 定数($CC_ : Class Constant)
	#-------------------------------------------------------------------------

	#-------------------------------------------------------------------------
	# 設定($CCONF_ : Class Config)
	#-------------------------------------------------------------------------
	# バッファサイズ
	[int] $CCONF_BufferSize = 1024

	# タイムアウト
	[int] $CCONF_TimeOut = 30

	# ログファイルフルパス
	[string] $CCONF_LogPath = ""

	# ログを記録しない
	[bool] $CCONF_NotRecordLog = $false



	#-------------------------------------------------------------------------
	# 変数($CV_ : Class Variable)
	#-------------------------------------------------------------------------
	# ソケット
	[System.Net.Sockets.TcpClient] $CV_Socket

	# ストリーム
	[System.Net.Sockets.NetworkStream] $CV_Stream

	# ライター
	[System.IO.StreamWriter] $CV_Writer

	# 受信バッファ
	[Byte[]] $CV_ReceiveBuffer

	# 送信バッファ
	[Byte[]] $CV_SendBuffer

	# 受信バッファ Index
	[int] $CV_ReceiveBufferIndex

	# 送信バッファ Index
	[int] $CV_SendBufferIndex

	#-------------------------------------------------------------------------
	# 内部 メソッド (protected 扱い)
	#-------------------------------------------------------------------------

	#-------------------------------------------------------------------------
	# 公開 メソッド (public 扱い)
	#-------------------------------------------------------------------------

	##########################################################################
	# コンストラクタ
	##########################################################################
	TCPClient(){
	}

	##########################################################################
	# 環境設定変更
	##########################################################################
	[void]SetEnvironment( [string]$LogPath, [bool]$NotRecordLog, [int] $TimeOut ){

		# ログパス
		if( $LogPath -ne [string]$null ){
			$this.CCONF_LogPath = $LogPath
		}

		$this.CCONF_NotRecordLog = $NotRecordLog

		# タイムアウト
		if( $TimeOut -ne [int]$null ){
			$this.CCONF_TimeOut = $TimeOut
		}
	}



	##########################################################################
	# ログ出力(public)
	##########################################################################
	[string] Log( [string]$LogString ){

		$Now = Get-Date

		$Log = $Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " "
		$Log += $LogString

		# ログフォルダーがなかったら作成
		$LogPath = Split-Path -Parent $this.CCONF_LogPath
		if( -not (Test-Path $LogPath) ) {
			New-Item $LogPath -Type Directory
		}

		if( -not $this.CCONF_NotRecordLog ){
			Write-Output $Log | Out-File -FilePath $this.CCONF_LogPath -Encoding utf8 -append
		}

		Return $Log
	}



	##########################################################################
	# 接続
	##########################################################################
	[void] Connect([string]$RemoteHost, [int]$Port ){

		# ログパス
		if( $this.CCONF_LogPath -eq "" ){
			$Now = Get-Date
			$YYYYMMDD = $Now.ToString("yyyy-MMdd")
			$this.CCONF_LogPath = Join-Path $PSScriptRoot "TCP_$YYYYMMDD.log"
		}

		# 受信バッファ
		$this.CV_ReceiveBuffer = New-Object byte[] $this.CCONF_BufferSize

		# 送信バッファ
		$this.CV_SendBuffer = New-Object byte[] $this.CCONF_BufferSize

		# Text バッファ
		# $this.CV_TextBuffer = New-Object byte[] $this.CCONF_BufferSize

		# tcp 接続
		Add-Type -AssemblyName System.Net

		$this.CV_Socket = New-Object System.Net.Sockets.TcpClient($RemoteHost, $Port)
		$this.CV_Stream = $this.CV_Socket.GetStream()
		$this.CV_Writer = New-Object System.IO.StreamWriter($this.CV_Stream)

	}


	##########################################################################
	# 切断
	##########################################################################
	[void]DisConnect(){
		$this.CV_Stream.Close()
		$this.CV_Writer.Close()
		$this.CV_Socket.Close()

		$this.CV_Stream.Dispose()
		$this.CV_Writer.Dispose()
		$this.CV_Socket.Dispose()
	}

	##########################################################################
	# 送信
	##########################################################################
	[void]Send( [string]$Message, [bool]$Display ){
		if( $Display -eq $true ){
			Write-Host $this.Log("[SEND] $Message")
		}
		else{
			Write-Host $this.Log("[SEND] ********")
		}


		$this.CV_Writer.WriteLine($Message)
		$this.CV_Writer.Flush()
	}

	##########################################################################
	# 受信
	##########################################################################
	[string[]] Receive ( [string]$Prompt ){

		# Text バッファ
		$TextBuffer = New-Object byte[] $this.CCONF_BufferSize

		# テキスト メッセージ
		[string]$TextMeaage = ""

		# エンコード
		$Encoding = New-Object System.Text.AsciiEncoding

		# 初期タイムアウト設定
		$Now = Get-Date
		$TimeOver = $Now.AddSeconds($this.CCONF_TimeOut)


		while( (Get-Date) -le $TimeOver ){

			# 少し待つ
			sleep 1

			# 送信バッファクリア
			$this.CV_SendBufferIndex = 0

			if( $this.CV_Stream.DataAvailable ){
				# 受信バッファ読み取り
				$Read = $this.CV_Stream.Read($this.CV_ReceiveBuffer, 0, $this.CCONF_BufferSize)

				# 受信したのでタイムアウト リセット
				$Now = Get-Date
				$TimeOver = $Now.AddSeconds($this.CCONF_TimeOut)

				# 受信バッファ解析
				for($this.CV_ReceiveBufferIndex = 0; $this.CV_ReceiveBufferIndex -lt $Read; ){
					# バイナリハンドリングするときの処理
					# if( $this.CV_ReceiveBuffer[$this.CV_ReceiveBufferIndex] -eq $this.CC_IAC ){
					#	# IAC 受信処理
					#	$this.ReceiveIAC()
					#}
					#else{
						# テキスト受信
						for( $i = 0; $this.CV_ReceiveBufferIndex -lt $Read; $i++ ){
							$TextBuffer[$i] = $this.CV_ReceiveBuffer[$this.CV_ReceiveBufferIndex++]
						}
						$RaceveTempText = ($Encoding.GetString( $TextBuffer, 0, $i ))
						$TextMeaage += $RaceveTempText
					#}
				}

				# 受信テキストがプロンプトだったら抜ける
				[string]$TmpBuffer = $TextMeaage -replace "`r",""
				[array]$Lines = $TmpBuffer -split "`n"
				if( $Lines.Count -ne 0 ){
					# 値の入った最終行を取得
					 $LastLine = ""
					for( $i = $Lines.Count; $i -gt 0; $i-- ){
						$LastLine = $Lines[$i -1]
						if( $LastLine.Length -ne 0 ){
							break
						}
					}

					if( $LastLine -Match $Prompt ){
						if( $this.CCONF_Debug ){ Write-Host $this.Log( "[DEBUG] 受信テキストがプロンプトに一致した : `"$Prompt`" / `"$LastLine`""  ) }
						if( -not $this.CV_Stream.DataAvailable ){
							break
						}
					}
					else{
						if( $this.CCONF_Debug ){ Write-Host $this.Log( "[DEBUG] 受信テキストがプロンプトに一致しない : `"$Prompt`" / `"$LastLine`"" ) }
					}
				}
				else{
					if( $this.CCONF_Debug ){ Write-Host $this.Log( "[DEBUG] 受信テキストサイズ Zero" ) }
				}
			}

			## ハンドシェーク等のバイナリコントロールする場合
			#if( $this.CV_SendBufferIndex -ne 0 ){
			#	$this.CV_Stream.Write($this.CV_SendBuffer, 0, $this.CV_SendBufferIndex)
			#	if( $this.CCONF_Debug ){ Write-Host $this.Log("[DEBUG] ---- Send Message ----") }
			#}
			#else{
			#	if( $this.CCONF_Debug ){ Write-Host $this.Log("[DEBUG] ---- Receive wait ----") }
			#}
			#
			## 送信後少し待つ
			#sleep 1
		}

		# テキスト受信していたらログに書く
		[string[]] $Lines
		if( $TextMeaage.Length -ne 0 ){
			$TmpBuffer = $TextMeaage -replace "`r",""
			$Lines = $TmpBuffer -split "`n"
			foreach( $Line in $Lines ){
				Write-Host $this.Log("[Receive] $Line")
			}
		}

		return $Lines
	}
}

##############################################
# POP3 Client Class
##############################################
class POP3 : TCPClient {

	#-------------------------------------------------------------------------
	# 定数($CC_ : Class Constant)
	#-------------------------------------------------------------------------

	#-------------------------------------------------------------------------
	# 設定($CCONF_ : Class Config)
	#-------------------------------------------------------------------------

	#-------------------------------------------------------------------------
	# 変数($CV_ : Class Variable)
	#-------------------------------------------------------------------------


	#-------------------------------------------------------------------------
	# 内部 メソッド (protected 扱い)
	#-------------------------------------------------------------------------
	##########################################################################
	# MD5 ハッシュ
	##########################################################################

	[string] GetMD5Hash( [string] $BaseString ){
		# バイト配列にする
		$ByteString = [System.Text.Encoding]::UTF8.GetBytes($BaseString)

		# アセンブリロード
		Add-Type -AssemblyName System.Security

		# MD5 オブジェクトの生成
		$MD5 = New-Object System.Security.Cryptography.MD5CryptoServiceProvider

		# Hash 値を求める
		$HashBytes = $MD5.ComputeHash($ByteString)

		# MD5 オブジェクトの破棄
		$MD5.Dispose()

		# Hash 値を16進文字列にする
		[string] $HashString = ""
		foreach( $HashByte in $HashBytes ){
			$HashString += $HashByte.ToString("x2")
		}

		return $HashString
	}




	#-------------------------------------------------------------------------
	# 公開 メソッド (public 扱い)
	#-------------------------------------------------------------------------

	##########################################################################
	# ログイン
	##########################################################################
	[void] Login(
		[string]$Server,
		[int]$Port,
		[string]$ID,
		[string]$Password
	){
		# 接続
		([TCPClient]$this).Connect( $Server, $Port )

		# 受信
		[string]$ReceiveData = ([TCPClient]$this).Receive(">")

		# チャレンジ生成
		$TmpBuffer = ($ReceiveData.Split("<"))[1]
		$TmpBuffer = ($TmpBuffer.Split(">"))[0]

		[string]$ChallengeBase = "<" + $TmpBuffer + ">" + $Password

		[string]$Challenge = $this.GetMD5Hash( $ChallengeBase )

		# APOP 認証

		$SendMessage = "APOP $ID $Challenge"
		([TCPClient]$this).Send( $SendMessage, $true )

		# 受信
		$ReceiveData = ([TCPClient]$this).Receive(".")

		return
	}

	##########################################################################
	# ログオフ
	##########################################################################
	[void] Logoff(){
		# Logoff コマンド送信
		([TCPClient]$this).Send( "QUIT", $true )

		# 受信
		$ReceiveData = ([TCPClient]$this).Receive(".")

		# 切断
		([TCPClient]$this).DisConnect()
	}

	##########################################################################
	# メールリスト取得
	##########################################################################
	[string[]] GetMessageList(){
		# コマンド送信
		([TCPClient]$this).Send( "LIST", $true )

		# 受信
		[string[]]$ReceiveData = ([TCPClient]$this).Receive(".")

		[String[]]$ReturnDatas = @()

		$Max = $ReceiveData.Count

		for($i=1; $i -lt $Max; $i++ ){
			$LineData = $ReceiveData[$i]
			$ListDatas = $LineData.Split(" ")
			if( $ListDatas.Count -eq 2 ){
				$ReturnDatas += $ListDatas[0]
			}
		}

		return $ReturnDatas
	}


	##########################################################################
	# メール受信
	##########################################################################
	[string[]] ReceiveMessage([string] $MessageNummber){
		# コマンド送信
		$Sendcommand = "RETR $MessageNummber"
		([TCPClient]$this).Send( $Sendcommand, $true )

		# 受信
		[string[]]$ReceiveData = ([TCPClient]$this).Receive(".")

		$Max = $ReceiveData.Count
		[string[]] $ReturnData = @()
		for($i =1; $i -lt $Max -2; $i++){
			$ReturnData += $ReceiveData[$i]
		}
		return $ReturnData
	}


	##########################################################################
	# メール削除
	##########################################################################
	[void] RemoveMessage([string] $MessageNummber){
		# コマンド送信
		$Sendcommand = "DELE $MessageNummber"
		([TCPClient]$this).Send( $Sendcommand, $true )

		# 受信
		[string[]]$ReceiveData = ([TCPClient]$this).Receive(".")

		return
	}

}




###########################################
# test main
###########################################

$POP3 = New-Object POP3

$POP3.Login( "MailServer", 110, "ID", "Password" )

$MessageNummbers = $POP3.GetMessageList()

$Message = $POP3.ReceiveMessage("1")

$Message

$POP3.RemoveMessage("1")

#foreach( $MessageNummber in $MessageNummbers ){
#	 $Message = $POP3.ReceiveMessage($MessageNummber)
#
#	 $Message
#}

$POP3.Logoff()

