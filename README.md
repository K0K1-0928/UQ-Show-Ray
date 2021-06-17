# UQ-Show-Ray

休日と休日の間に挟まる平日を見つけ、Google カレンダーに登録するプログラムです。  
コンテナバインド型の Google Apps Script ですので、  
以下のリンクからスプレッドシートをコピーして使用するのが簡単です。  
[UQ Show Ray](https://docs.google.com/spreadsheets/d/1EkURxni8GGe1sfkq1kwJ2V1TJmreZP4CHez9iuolHYM/edit?usp=sharing 'UQ Show Ray')

## How To Use

スプレッドシートをコピーして使用する事を前提とします。

1. [UQ Show Ray](https://docs.google.com/spreadsheets/d/1EkURxni8GGe1sfkq1kwJ2V1TJmreZP4CHez9iuolHYM/edit?usp=sharing 'UQ Show Ray')を開き、 ファイル > コピーを作成 でスプレッドシートをコピーします。
2. コピーしたスプレッドシートの設定値を書き換えます。
3. ツール > スクリプトエディタを開きます。
4. スクリプトエディタから、1 度 `UQShowRay()` を実行します。この際、スプレッドシートと Google カレンダーへの編集権限を要求されるので、許可します。
5. 以降は、トリガー設定により自動で実行されるようになります。

## スプレッドシートの設定値について

| 設定値           |    サンプル    | 説明                                                                                                |
| :--------------- | :------------: | :-------------------------------------------------------------------------------------------------- |
| 登録カレンダー   | xxxx@gmail.com | 処理実行時に登録する Google カレンダーの ID です。                                                  |
| 対象期間\_x ヶ月 |       6        | 処理を実行する際の対象期間を設定します。6 の場合、実行時から 6 ヶ月後までの範囲で処理を実行します。 |
| 通知設定\_x 日前 |       14       | 登録する休暇候補日のリマインダー設定です。14 を設定した場合、14 日前にリマインドします。            |
| 通知設定\_hour   |       9        | 登録する休暇候補日のリマインダー設定です。9 を設定した場合、午前 9 時にリマインドします。           |
| 通知設定\_minute |       30       | 登録する休暇候補日のリマインダー設定です。30 を設定した場合、x 時 30 分にリマインドします。         |
