function SoTGeo(fname)
%{
SoTGeo
DESCRIPTION
For measuring the internal condition of large, irregularly-shaped tree 
parts, this function creates a PiCUS geometry file (.pit) using 
measurements of the cross-sectional shape and nail positions obtained by 
any means at the measurement plane. Subsequently, this geometry file can 
be used to collect time of flight measurements using the PiCUS sonic 
tomograph. 

NOTES
The coordinates describing the cross-sectional shape and nail positions 
must be recorded in the same coordinate system. 

---------------------------------------------------------------------------
INPUTS
fname: string - filepath for Excel workbook containing geometry
measurements and required metadata

Specifically, the Excel workbook must be formatted as follows:
Sheet1: float - mx2 set of counter-clockwise ordered coordinates (cm) 
describing the cross-sectional shape of the measured tree part, where m >= 
99
Sheet2: float - mx2 set of counter-clockwise ordered coordinates describing 
the nail positions, where m == number of nails
Sheet3: float - 3x1 set of values describing the height of the measurement 
plane (cm), circumference at the height of measurement (mm), and the number
assigned to the nail oriented closest to North. 

OUTPUTS
Sheet4: float - new sheet written to Excel workbook containing 99x4 matrix 
showing the counter-clockwise ordered coordinates of the new shape in the 
first and second columns. In the third and fourth columns, the ID numbers 
of the new and original coordinates are displayed. 
PiCUS geometry file: a PiCUS geometry file will be created and stored in
the same directory as the input Excel workbook. The file name will be the
same as the workbook, except for a .pit file extension.
---------------------------------------------------------------------------
%} 

A=xlsread(fname,1); %Sheet1
B=xlsread(fname,2); %Sheet2
C=xlsread(fname,3); %Sheet3

%Translate into QI
B(:,1)=B(:,1)+(3-min(A(:,1)));
B(:,2)=B(:,2)+(3-min(A(:,2)));
A(:,1)=A(:,1)+(3-min(A(:,1)));
A(:,2)=A(:,2)+(3-min(A(:,2)));

%From set of n points, replace nearest neighbor in cartesian space
if size(A,1) < 99
    error('The length of A must be greater than or equal to 99');
elseif size(A,1) > 99
    D=[resamplePolyline(A,99), (1:99)', zeros(99,1)];
else
    D=[A, (1:99)', zeros(99,1)];
end
for k = 1:size(B,1)
    ids=setdiff(1:99,find(D(:,4)))';
    distances = sqrt((D(ids,1)-B(k,1)).^2+(D(ids,2)-B(k,2)).^2);
    [~,idx] = min(distances);
    D(ids(idx),1:2)=B(k,1:2);
    D(ids(idx),4)=k;
end

S=table(D(:,1),D(:,2),D(:,3),D(:,4),'VariableNames',{'X','Y','MP','NAIL'});
writetable(S,fname,'Sheet',4);

%Write PiCUS geometry .pit file
J{1,1}='';
J{2,1}='[Comments]';
J{3,1}='ort1=';
J{4,1}='ort2=';
J{5,1}='ort3=';
J{6,1}='ort4=';
J{7,1}='Baumnr=0';
J{8,1}='Formular=1';
J{9,1}='Baumart=';
J{10,1}='BaumartLatein=';
J{11,1}=strcat('Zeit=',datestr(datetime(),'mm/dd/yyyyHH:MM:SS AM'));
J{12,1}='StammUHoehe=130';
J{13,1}='StammU=';
J{14,1}='Baumhoehe=';
J{15,1}='KronenD=';
J{16,1}='Longitude=';
J{17,1}='Latitude=';
J{18,1}='KronenansatzHoehe=';
J{19,1}='Baumalter=';
J{20,1}='VitalitaetRoloff=0';
J{21,1}='Neigungsrichtung=';
J{22,1}='Neigungswinkel=0';
J{23,1}='Neigungbei=0';
J{24,1}='Bearbeiter1=';
J{25,1}='allg_kommentare1=';
J{26,1}='auftraggeber1=';
J{27,1}='BildDatei1=';
J{28,1}='BildDatei2=';
J{29,1}='';
J{30,1}='[Main]';
J{31,1}='Sensoranzahl=99';
J{32,1}='MiniSensorenanzahl=0';
J{33,1}='KlopfMethode=0';
J{34,1}='Hammer=0';
J(35:133,1)=sprintfc('ModTyp%-d=0',(1:99)');
J{134,1}='ModulVerstaerkung=0';
J{135,1}='SampleanzahlHuellkurve=40';
J{136,1}='gr_dm=0'; %Major diameter
J{137,1}='kl_dm=0'; %Minor diameter
J{138,1}='messPunktAbstand=0';
J{139,1}=strcat('u=',num2str(C(2))); %Girth
J{140,1}=strcat('Norden=',num2str(C(3))); %North at MP
J{141,1}='Pos1=N'; %Direction of MP 1
J{142,1}=strcat('Hoehe=',num2str(C(1))); %Height
J{143,1}='KDhomogen=-1';
J{144,1}='KDLoch=-1';
J{145,1}='KDRiss=-1';
J{146,1}='KDKern=-1';
J{147,1}='KDFaul=-1';
J{148,1}='KDPilz=-1';
J{149,1}='Hauptwindrichtung=1';
J{150,1}='';
J{151,1}='[NagelBorke]';
J{152,1}='NogelBorkeVerwenden=0';
J(153:251,1)=sprintfc('%-d=0/0',(1:99)');
J{252,1}='';
J{253,1}='[TreeSA]';
J{254,1}='Baumart=0';
J{255,1}='Druckfestigkeit=2';
J{256,1}='Luftwiderstandsbeiwert=0.25';
J{257,1}='Durchmesser1=0';
J{258,1}='Durchmesser2=0';
J{259,1}='Rindendicke=1';
J{260,1}='Baumhoehe=1';
J{261,1}='Standort=1';
J{262,1}='Kronenform=1';
J{263,1}='';
J{264,1}='[BPoints]';
J(265:363,1)=sprintfc('%-d=%-f/%-f',[(1:99)',D(:,1:2)]);
J{364,1}='';
J{365,1}='[ZBPoints]';
J(366:464,1)=sprintfc('%-d=%-d',[(1:99)',repmat(C(1),[99 1])]);
J{465,1}='';
J{466,1}='[MPoints]';
J(467:565,1)=sprintfc('%-d=%-f/%-f',[(1:99)',D(:,1:2)]);
J{566,1}='';
J{567,1}='[ZMPoints]';
J(568:666,1)=sprintfc('%-d=%-d',[(1:99)',repmat(C(1),[99 1])]);
J{667,1}='';
J{668,1}='[Diagnoses]';
J{669,1}='';

[fpath,name,~]=fileparts(fname);
fid=fopen(strcat(fpath,'\',name,'.pit'),'w');
fprintf(fid,'%s\r\n',J{:});
fclose(fid);

end
