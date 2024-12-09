% OpenDSSを使用して太田の潮流と電圧を模擬する
% 高圧配電線のモデルなどは修士論文を...
clear
Date = 20160828;
Dir = 'C:\Users\Sojun_Iwashina\program_temporary\simulation_of_flow\data\';
dir_output='C:\Users\Sojun_Iwashina\OneDrive - 東京理科大学\ドキュメント\卒研\program\flow_simulation\test\outputs\';
PVDir = 'D:\DATA\PV_output_1minute_data_set\'; %PV出力のフォルダ
LoadDir = 'D:\DATA\Demand_power_1minute_data_set\';%負荷データのフォルダ
NumNodes = 45;%44; 
NumHouses = NumNodes*3*4; %何軒読み込むか
period=24;

dt = datetime('now');
DateString = datestr(dt,'yyyyMMddHHmmssFFF');
%実行した日付と時刻のフォルダを作成
mkdir([dir_output,DateString])
dir_output = [dir_output,DateString,'\']; %結果の格納先を更新

% execute DSSStartup.m
[DSSStartOK, DSSObj, DSSText] = DSSStartup1;
% 回路読み込み
if DSSStartOK
    DSSText.command='Compile (C:\Users\Sojun_Iwashina\OneDrive - 東京理科大学\ドキュメント\卒研\program\flow_simulation\ota_simulation\Master.dss)';
end


% Set up the interface variables
DSSCircuit=DSSObj.ActiveCircuit;
DSSSolution=DSSCircuit.Solution;
DSSMon=DSSCircuit.Monitors;

PVpower=zeros(24,NumHouses);
PVpower_1min=zeros(24*60,NumHouses);
Load=PVpower;
Load_1min=zeros(24*60,NumHouses);
Demand=PVpower;
BatCharge = PVpower;
BatRemain = zeros(25,NumHouses);
BatCapacity=zeros(1,NumHouses);
BatInverter=zeros(1,NumHouses);
EVInverter=zeros(24,NumHouses);
EVCharge=zeros(24,NumHouses);
EVRemain = zeros(25,NumHouses);
Drivekyori=zeros(24,NumHouses);
Drivejikan=zeros(24,NumHouses);
Drivenenpi=zeros(24,NumHouses);
total_demand_per_hour=zeros(24,1);
PVpower=readmatrix([Dir,'Optimalflow',num2str(Date),'.xlsx'],'Sheet','Generate','Range','A1:TZ24');%元の範囲：A1:TN24
Load=readmatrix([Dir,'Optimalflow',num2str(Date),'.xlsx'],'Sheet','Load','Range','A1:TZ24');%元の範囲：A1:TN24
%{
BatCapacity=readmatrix([Dir,'Optimalflow',num2str(Date),'.xlsx'],'Sheet','BatCapacity','Range','A1:IT1');
EVInverter=readmatrix([Dir,'Optimalflow',num2str(Date),'.xlsx'],'Sheet','EVInverter','Range','A1:HT24');
BatInverter=readmatrix([Dir,'Optimalflow',num2str(Date),'.xlsx'],'Sheet','BatInverter','Range','A1:IT1');
Drivekyori=readmatrix([Dir,'Optimalflow',num2str(Date),'.xlsx'],'Sheet','DriveKyori','Range','A1:IT24');
Drivejikan=readmatrix([Dir,'Optimalflow',num2str(Date),'.xlsx'],'Sheet','Drivejikan','Range','A1:IT24');
Drivenenpi=readmatrix([Dir,'Optimalflow',num2str(Date),'.xlsx'],'Sheet','BatInverter','Range','A1:IT24');
%}

for i=1:24
    for j=1:NumHouses
        Demand(i,j)=Load(i,j)-PVpower(i,j);
    end
end

%{
%1時間値を線形補完で1分値に
timestamp=linspace(0,24,24*60);
hour=linspace(0,24,24);
for i=1:NumHouses
    Load_1min(:,i)=interp1(hour, Load(:,i), timestamp);
    PVpower_1min(:,i)=interp1(hour, Load(:,i), timestamp);
end
%}

DSSText.Command='set mode=daily loadmult=1 stepsize=1 number=1440';
% 負荷とPVデータをOpenDSSへ転送
for ii=1:NumHouses
    DSSText.command=['New Loadshape.Load',num2str(ii),...
        ' npts=1440 minterval=60 UseActual=true mult=(',num2str((Load(:,ii)/2).'), ')'];


    DSSText.command=['New Loadshape.PV',num2str(ii), ...
        ' npts=1440 minterval=60 mult=(',num2str(PVpower(:,ii).'),') useactual=true'];
end

% 低圧配電線作成
% 柱上変圧器作成 & 負荷PV接続
for ii=1:NumNodes
    % A相とB相
    %柱上変圧器作成
    DSSText.command=['New Transformer.LV_AB',num2str(ii),'Trans phases=1 Windings=3',...
    ' wdg=1 Bus=OH',num2str(ii),'.1.2 kV=6.6 kVA=30 Conn=LL',...
    ' wdg=2 Bus=LV_AB', num2str(ii), '.1.0 kV=0.10 kVA=15 Conn=LN',...
    ' wdg=3 Bus=LV_AB', num2str(ii), '.0.2 kV=0.10 kVA=15 Conn=LN'];

    %負荷とPV接続
    %　引き込み線
    DSSText.command=['New Line.Line_AB1_',num2str(ii),' bus1=','LV_AB', num2str(ii),...
        '.1.2',' bus2=BusLoad_AB',num2str(ii),'.1.2 phases=2 linecode=LV_12, units=km, length=1'];%低圧電線    
    DSSText.command=['New Line.Line_N',num2str(ii),' bus1=','LV_AB', num2str(ii), ...
        '.0',' bus2=BusLoad_AB',num2str(ii),'.3 phases=1 linecode=LV_N, units=km, length=1'];%中性線
    
    %低圧電線と中性線間に負荷を接続(負荷は100Vだけ)
    DSSText.command=['New Load.Load_AB',num2str(ii),'_1 bus1=BusLoad_AB',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+1),...
        '  mode=1 conn=LL status=varialbe  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AB',num2str(ii),'_2 bus1=BusLoad_AB',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+2),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AB',num2str(ii),'_3 bus1=BusLoad_AB',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+3),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AB',num2str(ii),'_4 bus1=BusLoad_AB',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+4),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    
    DSSText.command=['New Load.Load_AB',num2str(ii),'_1b bus1=BusLoad_AB',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+1),...
        '  mode=1 conn=LL status=varialbe  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AB',num2str(ii),'_2b bus1=BusLoad_AB',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+2),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AB',num2str(ii),'_3b bus1=BusLoad_AB',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+3),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AB',num2str(ii),'_4b bus1=BusLoad_AB',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+4),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];

    %200VにPVを接続
    DSSText.command=['New Generator.PV_AB',num2str(ii),'_1 bus1=BusLoad_AB',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+1),...
        ' mode=1 status=variable  Vminpu=0 Vmaxpu=5'];
    DSSText.command=['New Generator.PV_AB',num2str(ii),'_2 bus1=BusLoad_AB',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+2),...
        ' mode=1 status=variable  Vminpu=0 Vmaxpu=5'];
    DSSText.command=['New Generator.PV_AB',num2str(ii),'_3 bus1=BusLoad_AB',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+3),...
        ' mode=1 status=variable Vminpu=0 Vmaxpu=5'];
    DSSText.command=['New Generator.PV_AB',num2str(ii),'_4 bus1=BusLoad_AB',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+4),...
        ' mode=1 status=variable  Vminpu=0 Vmaxpu=5'];
    
    
%     %A相とC相
    DSSText.command=['New Transformer.LV_AC',num2str(ii),'Trans phases=1 Windings=3',...
    ' wdg=1 Bus=OH',num2str(ii),'.1.3 kV=6.6 kVA=30 Conn=LL',...
    ' wdg=2 Bus=LV_AC', num2str(ii), '.1.0 kV=0.10 kVA=15 Conn=LN',...
    ' wdg=3 Bus=LV_AC', num2str(ii), '.0.2 kV=0.10 kVA=15 Conn=LN'];

    DSSText.command=['New Line.Line_AC1_',num2str(ii),' bus1=','LV_AC', num2str(ii), '.1.2',' bus2=BusLoad_AC',...
        num2str(ii),'.1.2 phases=2 linecode=LV_12, units=km, length=1'];%低圧電線
    DSSText.command=['New Line.Line_N_AC',num2str(ii),' bus1=','LV_AC', num2str(ii), '.0',' bus2=BusLoad_AC',...
    num2str(ii),'.3 phases=1 linecode=LV_N, units=km, length=1'];%中性線
    %低圧電線と中性線間に負荷を接続
    DSSText.command=['New Load.Load_AC',num2str(ii),'_1 bus1=BusLoad_AC',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+5),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AC',num2str(ii),'_2 bus1=BusLoad_AC',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+6),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AC',num2str(ii),'_3 bus1=BusLoad_AC',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+7),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AC',num2str(ii),'_4 bus1=BusLoad_AC',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+8),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    
    DSSText.command=['New Load.Load_AC',num2str(ii),'_1b bus1=BusLoad_AC',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+5),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AC',num2str(ii),'_2b bus1=BusLoad_AC',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+6),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AC',num2str(ii),'_3b bus1=BusLoad_AC',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+7),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_AC',num2str(ii),'_4b bus1=BusLoad_AC',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+8),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    
    %200VにPVを接続
    DSSText.command=['New Generator.PV_AC',num2str(ii),'_1 bus1=BusLoad_AC',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+5),...
        ' mode=1 status=variable  Vminpu=0 Vmaxpu=5'];
    DSSText.command=['New Generator.PV_AC',num2str(ii),'_2 bus1=BusLoad_AC',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+6),...
        ' mode=1 status=variable  Vminpu=0 Vmaxpu=5'];
    DSSText.command=['New Generator.PV_AC',num2str(ii),'_3 bus1=BusLoad_AC',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+7),...
        ' mode=1 status=variable  Vminpu=0 Vmaxpu=5'];
    DSSText.command=['New Generator.PV_AC',num2str(ii),'_4 bus1=BusLoad_AC',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+8),...
        ' mode=1 status=variable  Vminpu=0 Vmaxpu=5'];
   
%     %B相とC相
    DSSText.command=['New Transformer.LV_BC',num2str(ii),'Trans phases=1 Windings=3',...
    ' wdg=1 Bus=OH',num2str(ii),'.2.3 kV=6.6 kVA=30 Conn=LL',...
    ' wdg=2 Bus=LV_BC', num2str(ii), '.1.0 kV=0.10 kVA=15 Conn=LN',...
    ' wdg=3 Bus=LV_BC', num2str(ii), '.0.2 kV=0.10 kVA=15 Conn=LN'];

    DSSText.command=['New Line.Line_BC1_',num2str(ii),' bus1=','LV_BC', num2str(ii), ...
        '.1.2',' bus2=BusLoad_BC',...
        num2str(ii),'.1.2 phases=2 linecode=LV_12, units=km, length=1'];%低圧電線
    DSSText.command=['New Line.Line_N_BC',num2str(ii),' bus1=','LV_BC', num2str(ii),...
        '.0',' bus2=BusLoad_BC',num2str(ii),'.3 phases=1 linecode=LV_N, units=km, length=1'];%中性線
    %低圧電線と中性線間に負荷を接続
    DSSText.command=['New Load.Load_BC',num2str(ii),'_1 bus1=BusLoad_BC',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+9),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_BC',num2str(ii),'_2 bus1=BusLoad_BC',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+10),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_BC',num2str(ii),'_3 bus1=BusLoad_BC',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+11),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_BC',num2str(ii),'_4 bus1=BusLoad_BC',num2str(ii),...
        '.1.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+12),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];

    DSSText.command=['New Load.Load_BC',num2str(ii),'_1b bus1=BusLoad_BC',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+9),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_BC',num2str(ii),'_2b bus1=BusLoad_BC',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+10),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_BC',num2str(ii),'_3b bus1=BusLoad_BC',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+11),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    DSSText.command=['New Load.Load_BC',num2str(ii),'_4b bus1=BusLoad_BC',num2str(ii),...
        '.2.3 phases=1  kV=0.10 kW=0 pf=1 Daily=Load',num2str((ii-1)*12+12),...
        '  mode=1 conn=LL status=variable  Vminpu=0.0 Vmaxpu=5'];
    
    %200VにPVを接続
    DSSText.command=['New Generator.PV_BC',num2str(ii),'_1 bus1=BusLoad_BC',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+9),...
        ' mode=1 status=variable  Vminpu=0 Vmaxpu=5'];
    DSSText.command=['New Generator.PV_BC',num2str(ii),'_2 bus1=BusLoad_BC',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+10),...
        ' mode=1 status=variable  Vminpu=0 Vmaxpu=5'];
    DSSText.command=['New Generator.PV_BC',num2str(ii),'_3 bus1=BusLoad_BC',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+11),...
        ' mode=1 status=variable  Vminpu=0 Vmaxpu=5'];
    DSSText.command=['New Generator.PV_BC',num2str(ii),'_4 bus1=BusLoad_BC',num2str(ii),...
        '.1.2 phases=1 conn=LL kV=0.20 kW=0 pf=1 Daily=PV',num2str((ii-1)*12+12),...
        ' mode=1 status=variable  Vminpu=0 Vmaxpu=5'];
    
    
    %モニター設置(結果確認用)
    % 電力メータ(A相負荷1，PV1，高圧配電線OH)
    % 電力メータ(A相負荷1，PV1)
        DSSText.command=['New Monitor.P_Load_AB',num2str(ii),' Load.Load_AB',num2str(ii),'_1 mode=1 PPolar=no terminal=1'];
        DSSText.command=['New Monitor.P_PV_AB',num2str(ii),' Generator.PV_AB',num2str(ii),'_1 mode=1 PPolar=no terminal=1'];
        %電圧メータ(A相負荷1，PV1)
        DSSText.command=['New Monitor.V_Load_AB',num2str(ii),' Load.Load_AB',num2str(ii),'_1 mode=0 terminal=1'];
        DSSText.command=['New Monitor.V_PV_AB',num2str(ii),' Generator.PV_AB',num2str(ii),'_1 mode=0 terminal=1'];
        % 電力メータ(B相負荷1，PV2)
        DSSText.command=['New Monitor.P_Load_AC',num2str(ii),' Load.Load_AC',num2str(ii),'_1 mode=1 PPolar=no terminal=1'];
        DSSText.command=['New Monitor.P_PV_AC',num2str(ii),' Generator.PV_AC',num2str(ii),'_1 mode=1 PPolar=no terminal=1'];
        %電圧メータ(B相負荷1，PV2)
        DSSText.command=['New Monitor.V_Load_AC',num2str(ii),' Load.Load_AC',num2str(ii),'_1 mode=0 terminal=1'];
        DSSText.command=['New Monitor.V_PV_AC',num2str(ii),' Generator.PV_AC',num2str(ii),'_1 mode=0 terminal=1'];
        % 電力メータ(C相負荷1，PV3)
        DSSText.command=['New Monitor.P_Load_BC',num2str(ii),' Load.Load_BC',num2str(ii),'_1 mode=1 PPolar=no terminal=1'];
        DSSText.command=['New Monitor.P_PV_BC',num2str(ii),' Generator.PV_BC',num2str(ii),'_1 mode=1 PPolar=no terminal=1'];
        %電圧メータ(C相負荷1，PV3)
        DSSText.command=['New Monitor.V_Load_BC',num2str(ii),' Load.Load_BC',num2str(ii),'_1 mode=0 terminal=1'];
        DSSText.command=['New Monitor.V_PV_BC',num2str(ii),' Generator.PV_BC',num2str(ii),'_1 mode=0 terminal=1'];
    DSSText.command=['New Monitor.P_OH',num2str(ii),' Line.OH',num2str(ii),' mode=1 PPolar=no terminal=1'];
    
   
%     if ii~=23
%         DSSText.command=['New Monitor.P_OH',num2str(ii),' Line.OH',num2str(ii),' mode=1 PPolar=no terminal=1'];
%     end
    
    DSSText.command=['New Monitor.V_OH',num2str(ii),' Line.OH',num2str(ii),' mode=0 terminal=1'];
%     if ii~=23
%         DSSText.command=['New Monitor.V_OH',num2str(ii),' Line.OH',num2str(ii),' mode=0 terminal=1'];
%     end
    
end

% ソース電力＆電圧
DSSText.command='New Monitor.P_all Line.OH1 mode=1 PPolar=no terminal=1';
DSSText.command='New Monitor.Vsource Vsource.source mode=0 terminal=1';

DSSText.command='Set VoltageBases = [6.6 0.10 0.20]';
DSSText.command='set controlmode=static';
DSSText.Command='set mode=daily stepsize=60.0 number=1440';
DSSSolution.Solve;

%% プロット
Time=linspace(0,24,1440);
% ソースの供給電力
DSSMon.name='P_all';
P(1:3,:) = ExtractMonitorData(DSSMon,1:2:5,1.0); %A,B,C相潮流
figure(1);
plot(Time,sum(P,1));
%{
hold on
plot(Time,-sum(Load(:,1:NumHouses),2)-sum(PVpower(:,1:NumHouses),2),'--'); hold off
%}
xlim([0 24]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
'FontWeight','Bold')
xlabel('Time [hour]'); ylabel('Power Flow [kW]'); grid on; %legend('Simulated', 'Load - PV')
saveas(gcf, dir_output + "power_from_source.png");
saveas(gcf, dir_output + "power_from_source.fig");
%disp('done');

% 電圧変化のプロット
%高圧系統
for ii=1:NumNodes
    DSSMon.name=['V_OH',(num2str(ii))];
    OH(ii).V(1:6,:) = ExtractMonitorData(DSSMon,1:6,1.0);
end

figure(2);
for ii=1:NumNodes%[1:10:41,NumNodes]
    plot(Time, OH(ii).V(1,:)*sqrt(3));hold on;
end, hold off
xlim([0 24]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
'FontWeight','Bold')
xlabel('Time [hour]'); ylabel('Line Voltage [V]'); grid on; 
saveas(gcf, dir_output + "voltage_high_voltage_grid.png");
saveas(gcf, dir_output + "voltage_high_voltage_grid.fig");
%disp('done');

for ii=1:NumNodes %ii=[1:10:41,45]
    h = figure('visible','off');
    plot(Time, OH(ii).V(1,:)*sqrt(3));
    xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
    'FontWeight','Bold')
    xlabel('Time [hour]'); ylabel('Line Voltage [V]'); grid on;
    saveas(gcf, dir_output + "voltage_high_voltage_grid_node_" + num2str(ii) + ".png");
    saveas(gcf, dir_output + "voltage_high_voltage_grid_node_" + num2str(ii) + ".fig");
end

%低圧系統
for ii=1:NumNodes
    DSSMon.name=['V_Load_AB',(num2str(ii))];
    Load_AB(ii).V(1:6,:) = ExtractMonitorData(DSSMon,1:6,1.0);
end

for ii=1:NumNodes
    DSSMon.name=['V_Load_AC',(num2str(ii))];
    Load_AC(ii).V(1:6,:) = ExtractMonitorData(DSSMon,1:6,1.0);
end

for ii=1:NumNodes
    DSSMon.name=['V_Load_BC',(num2str(ii))];
    Load_BC(ii).V(1:6,:) = ExtractMonitorData(DSSMon,1:6,1.0);
end

figure(3);
for ii=1:NumNodes %ii=[1:10:41,45]
    plot(Time, Load_AB(ii).V(1,:));hold on;
    plot(Time, Load_AC(ii).V(1,:));hold on;
    plot(Time, Load_BC(ii).V(1,:));hold on;
end, hold off
xlim([0 24]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
'FontWeight','Bold')
xlabel('Time [hour]'); ylabel('Line Voltage [V]'); grid on;
saveas(gcf, dir_output + "voltage_low_voltage_grid.png");
saveas(gcf, dir_output + "voltage_low_voltage_grid.fig");


for ii=1:NumNodes %ii=[1:10:41,45]
    h = figure('visible','off');
    plot(Time, Load_AB(ii).V(1,:));
    xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
    'FontWeight','Bold')
    xlabel('Time [hour]'); ylabel('Line Voltage [V]'); grid on;
    saveas(gcf, dir_output + "voltage_low_voltage_grid_node_" + num2str(ii) + ".png");
    saveas(gcf, dir_output + "voltage_low_voltage_grid_node_" + num2str(ii) + ".fig");

    h = figure('visible','off');
    plot(Time, Load_AC(ii).V(1,:));
    xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
    'FontWeight','Bold')
    xlabel('Time [hour]'); ylabel('Line Voltage [V]'); grid on;
    saveas(gcf, dir_output + "voltage_low_voltage_grid_node_" + num2str(ii) + ".png");
    saveas(gcf, dir_output + "voltage_low_voltage_grid_node_" + num2str(ii) + ".fig");

    h = figure('visible','off');
    plot(Time, Load_BC(ii).V(1,:));
    xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
    'FontWeight','Bold')
    xlabel('Time [hour]'); ylabel('Line Voltage [V]'); grid on;
    saveas(gcf, dir_output + "voltage_low_voltage_grid_node_" + num2str(ii) + ".png");
    saveas(gcf, dir_output + "voltage_low_voltage_grid_node_" + num2str(ii) + ".fig");
end

%潮流のプロット
for ii=1:NumNodes
    DSSMon.name=['P_OH',(num2str(ii))];
    OH(ii).P(1:6,:) = ExtractMonitorData(DSSMon,1:6,1.0);
    %disp(size(OH(ii).P(1:6,:)));
end
figure(4);
for ii=1:NumNodes %ii=[1:10:41,45]
    plot(Time, OH(ii).P(1,:));hold on;
end, hold off
xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
'FontWeight','Bold')
xlabel('Time [hour]'); ylabel('Power [kW]'); grid on;
saveas(gca, dir_output + "flow_high_voltage_grid.png"); %saveas(gcf, dir_output + "flow_high_voltage_grid.png");
saveas(gca, dir_output + "flow_high_voltage_grid.fig");

for ii=1:NumNodes %ii=[1:10:41,45]
    h = figure('visible','off');
    plot(Time, OH(ii).P(1,:));
    xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
    'FontWeight','Bold')
    xlabel('Time [hour]'); ylabel('Power [kW]'); grid on;
    saveas(gcf, dir_output + "flow_high_voltage_grid_node_" + num2str(ii) + ".png");
    saveas(gcf, dir_output + "flow_high_voltage_grid_node_" + num2str(ii) + ".fig");
end

%低圧系統の各相における潮流
for ii=1:NumNodes
    % 呼び出し箇所の修正
    [Load_AB(ii).P(1:6, :), y_size] = ExtractMonitorData(DSSMon, 1:6, 1.0);

    % 配列の大きさを確認して範囲外エラーを防ぐ
    if max(1:6) + 2 > y_size(1)
        error('指定されたチャンネル番号が配列yの範囲を超えています．');
    end
end

for ii=1:NumNodes
    DSSMon.name=['P_Load_AC',(num2str(ii))];
    Load_AC(ii).P(1:6,:) = ExtractMonitorData(DSSMon,1:6,1.0);
end

for ii=1:NumNodes
    DSSMon.name=['P_Load_BC',(num2str(ii))];
    Load_BC(ii).P(1:6,:) = ExtractMonitorData(DSSMon,1:6,1.0);
end

figure(6);
for ii=1:NumNodes %ii=[1:10:41,45]
    plot(Time, Load_AB.P(1,:));hold on;
    plot(Time, Load_AC.P(1,:));hold on;
    plot(Time, Load_BC.P(1,:));hold on;
end, hold off
xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
'FontWeight','Bold')
xlabel('Time [hour]'); ylabel('Power [kW]'); grid on;
saveas(gca, dir_output + "flow_low_voltage_grid.png"); %saveas(gcf, dir_output + "flow_high_voltage_grid.png");
saveas(gca, dir_output + "flow_low_voltage_grid.fig");

for ii=1:NumNodes %ii=[1:10:41,45]
    h = figure('visible','off');
    plot(Time, Load_AB.P(1,:));
    xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
    'FontWeight','Bold')
    xlabel('Time [hour]'); ylabel('Power [kW]'); grid on;
    saveas(gcf, dir_output + "flow_low_voltage_grid_phase_A_node_" + num2str(ii) + ".png");
    saveas(gcf, dir_output + "flow_low_voltage_grid_phase_A_node_" + num2str(ii) + ".fig");

    h = figure('visible','off');
    plot(Time, Load_AC.P(1,:));
    xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
    'FontWeight','Bold')
    xlabel('Time [hour]'); ylabel('Power [kW]'); grid on;
    saveas(gcf, dir_output + "flow_low_voltage_grid_phase_B_node_" + num2str(ii) + ".png");
    saveas(gcf, dir_output + "flow_low_voltage_grid_phase_B_node_" + num2str(ii) + ".fig");

    h = figure('visible','off');
    plot(Time, Load_BC.P(1,:));
    xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
    'FontWeight','Bold')
    xlabel('Time [hour]'); ylabel('Power [kW]'); grid on;
    saveas(gcf, dir_output + "flow_low_voltage_grid_phase_C_node_" + num2str(ii) + ".png");
    saveas(gcf, dir_output + "flow_low_voltage_grid_phase_C_node_" + num2str(ii) + ".fig");
end
%}
Time=linspace(0,24,24);
figure(5)
for ii=1:NumNodes %ii=[1:10:41,45]
    plot(Time, Demand(:,ii));hold on;
end, hold off
xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
'FontWeight','Bold')
xlabel('Time [hour]'); ylabel('Demand [kW]'); grid on;
saveas(gca, dir_output + "demand.png"); %saveas(gcf, dir_output + "flow_high_voltage_grid.png");
saveas(gca, dir_output + "demand.fig");

for ii=1:NumNodes %ii=[1:10:41,45]
    h = figure('visible','off');
    plot(Time, Demand(:,ii));
    xlim([0 period]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
    'FontWeight','Bold')
    xlabel('Time [hour]'); ylabel('Power [kW]'); grid on;
    saveas(gcf, dir_output + "demand_house_" + num2str(ii) + ".png");
    saveas(gcf, dir_output + "demand_house_" + num2str(ii) + ".fig");
end

%{
%電圧不平衡？末端のOH13の各相の電圧をプロット
figure(3)
for ii=[1 3 5]
    plot(Time,OH(41).V(ii,:)*sqrt(3));hold on
end, hold off
xlim([0 24]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
'FontWeight','Bold')
xlabel('Time [hour]'); ylabel('Line Voltage [V]'); grid on; legend('A','B','C')

DSSMon.name='V_Load45';
V_Load=ExtractMonitorData(DSSMon,1:4,1.0);
V_abs=abs(V_Load(1,:).*exp(1i*V_Load(2,:)*pi/180)-V_Load(3,:).*exp(1i*V_Load(4,:)*pi/180));
figure(4);plot(Time, V_Load(1,:))
xlim([0 24]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
'FontWeight','Bold')
xlabel('Time [hour]'); ylabel('Load Voltage [V]'); grid on;

% SVRのタップをプロット
DSSMon.name='SVR_A';
SVR_tap(1,:)=ExtractMonitorData(DSSMon,1,1.0);
DSSMon.name='SVR_B';
SVR_tap(2,:)=ExtractMonitorData(DSSMon,1,1.0);
DSSMon.name='SVR_C';
SVR_tap(3,:)=ExtractMonitorData(DSSMon,1,1.0);
% 
figure(5); plot(Time,SVR_tap(:,1:end))
xlim([0 24]); set(gca, 'FontName', 'Helvetica', 'FontSize', 14, 'XTick',0:6:24,...
'FontWeight','Bold')
xlabel('Time [hour]'); ylabel('Tap [p.u.]'); ylim([0.98 1.02]); legend('A', 'B','C')
set(gca, 'YTickLabel',num2str(get(gca,'YTick').','%.3f'))

% % 高圧系統をプロット
%  DSSText.command='buscoords allnodecoords.csv';
%  DSSText.command='Set nodewidth=7';
%  DSSText.command='plot 1phLineStyle dots=y';
%}

DSSCircObj = actxserver('OpenDSSEngine.DSS');

disp('done');
%figure; plotCircuitLines(DSSCircObj)