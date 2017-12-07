◆ TCP Client と POP3 Class

これは、POP3 をハンドリングする PowerSehll Class です

L4 までの実装で、ユーザー認証とメッセージの受信と削除が出来ます
それより上位層はハンドリングしていません
認証は APOP のみ対応

POP3 ハンドリングを単純化するために、L3 ハンドリング用の TCP Client Class に分かれています

使い方は、test main を見てください

