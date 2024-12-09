# Distribution_System_test
## デバッグ中のプログラムです．デバッグ以外の目的で使わないでください．
## This program is being debugged. Please don't use it unless you wolud like to debug it.
#
説明
#
MATLABからOpenDSSを起動して配電系統のシミュレーションをするプログラムです。
低圧系統と高圧系統の電圧や潮流のグラフや，電源から流入する電力のグラフを描画することを想定しています．
低圧系統と高圧系統の電圧や潮流グラフは，すべてのノードや需要家におけるシミュレーション結果を1枚にまとめたものと，各ノードや各需要家ごとのグラフが出力される想定です．
#
使用方法
1. 現在リポジトリにあるフォルダ・ファイルを同じディレクトリに保存してください．
2. Ota_simulation_oneday.mを実行してください．
#
バグの内容
実行すると，次のエラーメッセージが出力されます．
#
位置 1 のインデックスが配列範囲を超えています。インデックスは 6 を超えてはなりません。

エラー: ExtractMonitorData (行 18)
#
        sa = y(chan+2,:) / base;
#
エラー: Ota_simulation_oneday (行 432)
#
    Load_AC(ii).P(1:6,:) = ExtractMonitorData(DSSMon,1:6,1.0);
#
すでに試したこと
#
コマンドウィンドウから
>> sa = y(chan+2,:) / base
#
>> y(chan+2,:)
#
>> chan
#
>> size(y)
#
を実行しましたが，「関数または変数 'y' が認識されません。」というメッセージが出力されます．
