<%
Dim Startime,SqlNowString,ConnStr
Dim conn,Db,MyDbPath,pos
Startime = Timer()
MyDbPath = ""
MyDbPath=request.ServerVariables("url")
pos=instr(MyDbPath,"/")+1
MyDbPath=left(MyDbPath,instr(pos,MyDbPath,"/"))


Const IsSqlDataBase = 0


If IsSqlDataBase = 1 Then
	Const SqlDatabaseName = "dzwangweb2"
	Const SqlPassword = ""
	Const SqlUsername = "sa"
	Const SqlLocalName = "(local)"
	SqlNowString = "GetDate()"
Else
	Db = "/database/#Website_data.mdb"
	SqlNowString = "Now()"
End If

	If IsSqlDataBase = 1 Then
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	Else
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(Db)
	End If
    On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open ConnStr
	If Err Then
		err.Clear
		Set Conn = Nothing
		Response.Write "���ݿ����ӳ���,���������ִ�"
		Response.End
	End If

sub CloseConn()

end sub


sub htmlend
response.write "<p>"
response.write "<table cellspacing=0 cellpadding=0 width=95% align=center><tr>"
response.write "<td align=middle> Copyright 2008 KunNet <br>"
response.write "Powered by  <a href=""#"" target=_blank ><font color=""#FF0000"">ȫѶ���ٷ���վ</font></a>&nbsp;&copy; 2008<br>"
response.write "</td></tr></table>"

end sub

function showinfo(place)

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from [siteinfo] where place='"&place&"' "
	rs.Open sql,conn,1,3
	if not rs.eof then 
		if rs("open")=1 then
		showinfo=rs("content")
		end if
	end if
	rs.Close
end function

function htmltojs(str)
if isnull(str) then exit function
 dim tojs,i
 
 str = Replace(str, Chr(10), "") 
 str = Replace(str, Chr(32) & Chr(32), "") 
 str = Split(str, Chr(13))
 tojs=""
 For i = 0 To UBound(str) 
 If Trim(str(i)) <> "" Then 
 str(i) = Replace(str(i), Chr(34), Chr(39)) 
 'tojs = tojs & "document.write(" & Chr(34) & str(i) & Chr(34) & ");" & Chr(10) 
 tojs = tojs & str(i)
 End If 
 Next
 htmltojs = tojs
end function


function BigToGb(content)
 dim s,t,c,d,i
s="����רҵ�Զ�˿������ɥ����Ϊ������������ϰ������������ب�ǲ�Ķ�����ڽ����ز����Ǽ����Ż���ɡΰ����������α���������½�����ȿ�٭ٯٶٱٲ��ٳ��ծ�������ǳ����ϴ��ж��������������������ڸԲ�д��ũ�������������������ݷ���ƾ������ۻ����մ�ɾ����ɲ�����ܼ��н�����Ȱ����۽����������ѫ��������ҽ��Э����¬±����ȴ�᳧������ѹ���ǲ������ó���������˫�������Ҷ��̾ߴ���������������߼߽Ż߿��Ա��Ǻ��ӽ���������������ܻ�������Ӵ�����﻽��������Х�������������������԰��Χ���ͼԲʥ�۳������̳�ް����׹¢���ݿ��ѵ�����������ǵ��ǽ׳���Ǻ���������ͷ�ж���ۼ�ܽ���ױ����������¦�欽������Ӥ���������������ѧ������ʵ�����ܹ�����޶�Ѱ���ٽ�������Ң�Ͼ������������������᫸��ᰵ�����Ͽ��������ո�������۹��ϱ�˧ʦ�����Ĵ�֡���������ݹ���®�п�Ӧ���ӷ��޿�����߱�������䵯ǿ�鵱¼�峹�����������黳̬���������������������Ҷ������������������������ҳͱ�㫲ѵ���㳷�����������Ϸ�ս꯻�ִ����ɨ���Ÿ����ҿ�������������£��ӵ��š����ֿ����̢Ю�ӵ�������������񻻵���°����������������§��Я�����ҡ��̯������ߢߣߥ���ܵ�����ի�ն���޾�ʱ������Խ�����������������ɱ��Ȩ������迼�������������ǹ����������դ��ջ���ж��������������������嵵��������׮�μ�����������¥���鷼����ƺ���ӣ�ͳ������ݻ��ŷ���������������Ź��챱ϱ�ձ���������ۻ㺺���ڹ�û��Ž���ٲ���Ţ���������к�������������ǳ������ǲ�䫼�䯻��Ũ����������л�������ɬ��Ԩ�����½�������������ʪ������������������������б�̲������ΫǱ���������������ֲ��¯�����������˸�������̷����ǻ��̽��Ȼ����Ⱞү�ǣ����״���̱�������ʨ����������������è��̡�����⻷�����巩���������������걵续�������ű���񴯷�Ӹ�����컾���������̱���Ѣ񲰨����յ�μ�ǵ�������������غ�������������ש�����������˶����ȷ�ﰭ���׼���������������»����ͺ�����ػ��ƻ�˰�����������Ҥ���ѿ��������������ȼ�������ɸ�ݳ�ǩ����������������¨�����������������������ֽ��������������Լ����������γ硴��ɴ�������ڷ�ֽ�Ʒ�Ŧ������������ϸ֯��称������ﾭ窰��޽���笻��Ѥ������ͳ�篾������м�簼�������糴����ά��緱������������׺��罼������翼�����綶������Ļ����Ʊ���Ե�Ƹ����Ƿ��ɲ���������������ӧ�������������������ٽ�������޷��������������ʳ�����ְ���������೦����������в�е�ʤ�����ֽ�����������ŧ�����������������������������������ս���ܼ��«��έ�����ɲ�����ƻ����������뾣���������������ٻ�����ӫݡݣݥ��ݦҩݰ����ݪݫݲ��ݵӨݺ��өӪ�����������޽���������������Ǿ������ޭ��޴޺²�������Ϻ�ʴ�����������������������������Ӭ���Ы��������β����������Ϯװ�����Ͽ������ܼ��۹����������������������������������ܼƶ����ϼ�ڦڧ����ڨ��ѵ��Ѷ�ǽ���کڪ��ګ������Ϸ���þ�֤ڬڭ����ʶթ����ڮ�ߴ�ڰگ��ڱڲڳ��ڴʫڵڶ����ڷ����ڸڹ��ѯ��ں�����ڻڼ������ڽ��ھ�ջ�ڿ˵����������ŵ���·̿�����˭�ŵ�����׻��̸��ı�ȵ�����г����ν�����β�������������������лҥ����ǫ�׽�á������̷������������Ǵ���߱��긺�������Ͱ��˻��ʷ�̰ƶ�Ṻ���ᷡ�����������ܴ�ó�Ѻ������޼ֻ�����¸���������������޸������ʹ���������׸��׬������������Ӯ���Ը�������Ծ���ȼ���������ӻ�������������������������������ת��������������������������������������������������Թ�����ꣷ�������ԯϽշ���ꥴǱ����ɴ�Ǩ�����˻����ԶΥ�������ɼ���ѡѷ��������ң��������������ۣۦ֣۩۪�ǵ��ͽ��������ͼ����������붤�����������̷��������θ����Ҷ۳����Ʊ�������Կ�վ��ٹ���������ť����Ǯ��ǯ�ܲ����������������������������Ǧí������������������������������ͭ������ա��ϳ���������������綠�ҿ����������������������﮳��������ﲷ�п������������ê��������സ׶�����Ķ�����������������������������̶�þ�����������������ָ���������F���������޾������������������������������������ⳤ�������Ʊ��ʴ��������ȼ�������բ�ֹ��������̷���������������������������ղ����������������۶�������׼�½¤���������������������ѳ���������������������Τ�ͺ������ҳ��������˳������˶���������Ԥ­���ľ��������Ƶ���ӱ�������ն��������ȧ������Ʈ쭷����м����������������α����¶����ýȱ��Ķ��������ڹ����Ȳ���������������Ԧ��ѱ������¿��ʻ�������פ���������������溧������鿥����������ƭ��ɧ���������������������������������������³�������ֱ���������������������������������������������������������������������������������������������������������������𯼦���Ÿѻ�����Ѽ����ԧ�������������������������ȵ�������������������Ϻ�����������������ӥ���������������������������촳�������������������ȣ���������"
 t="�f�c���I���|�z�G�ɇ��ʂ��R�����e�x���������l���I�y��̝�����a���H�C�|�H�ā��}�x���r���������ゥ�����t����΁��w�L�b�H�e�ɂȃS�~���z�R�����z�������A��E�f�����������������h�m�P�dƝ�B�F�σȌ��Ԍ�܊�r�T�_�Q�r���Q�D���p���C�P�D�{�P����c���t�����h�e�q�x�������������������k�Մ�ӄ�ńڄ݄ׄ�Q�T�^�t�A�f���u�R�u�P�l�s���S�d�v���������������B�N���h�����p�l׃���B�~̖�@�\���Άᇍ ���ǅȇ`�҇I���h�T�J�܆�ԁ����푆��}�^�􇂇W�������чO�Z��K�݇ʇ��[���D���ˇ��u�ڇ��̈F�@��������D�A�}�����ĉK�ԉ��ȉΉ]�����ŉ������׉|���N�P�_��q����������̎��͉��^�A�Z�Y�J�^���W�y�D������������I�ƋɌD�ʋz����ȋ����܋�ԋߌO�W�\�������������m���e���������ی����m�L����M�ӌόÌٌҌՎZ�q�M�獏�s���u�X�h�F�{�����n������V����p��^�Ŏ��������Î����͎Ύ�����V�c�]�T�쑪�R���U�[�_�����s�������������w��䛏��؏��Ƒ��ԑn���ёB�Z���Y����z�������ِa�����Q�Ð�Ő�������ґa���@�֑K�͑v�ܑM���T�C���|�ؑ��Б��ߑ������̔U�ВߓP�_�ᒁ���������o����M�n����r�Q�ܓ񓴔���钶�ϓ��ג�D�]�Ɠp��Q�v����S�ۓ�������v�R�����y�z�d�[�u�P���t�Δf�X�]�x�\���������S�̔ؔ��o�f�r��ҕ��@�x�ԕϕ�������X�C���s���l���q�O�����З����������n���f�d�Ř˗����ɗ������ژ䗫�ә藿��E�n�����u�������z�И�������E�Ǚ���Ι������M�{�љ��������_�g�e�W���{����������������ݞ�����֚Кښ�������R�h�����ϛ]���a�r�S�朿�����I�͞{�o�T�a���ɛܝ����D�ќ\�{������y�ҝ��g���G�❡�����Z�i���u�o�읙���q���՜Y�O�n�^�u�ƝO�c�B�؞��񝢞R�s���L�������M�]�V�E���I���u�t���H���z���|�l������`�ĠN���t�������c����q���N�T��������Z�C�a�៨�F�c�۠������ޠ٠�E�q�N�����M�{���b�z�s���C�J�M�i؈�I�H�^���|�h�F�t�k�m���q�I�������a�����T늮��������X�����O�������b�d�W�A���B�w�D���T�c�a�`�]�_�d�}�����K�}�O�w�I�P�[�����A���G�m���C���\�V�X�a�u�����Z�a�[�A�T�����_�|�K���~�|�R�Y�����\���A���U�x�d���N�z�e�Q�x���d���w�F�`�[�G�Z�C�Q�]�M�Q���V�S�P�a�{�\�e�`�Y�~�I�����j�D�X�j�������t�@�h�f�[�eg�c���S�Z�R�o�{�m�u�t�q�w�v�s���w�k�o�x�����������V�{�v�]�����y���~�����C�������M�������K�U�O�E�I�B�[���H���q�Y�@�W�L�o�k�{�j�^�g�y�������C���d�^�����w�c�m�_�p�b�i�K�S�d�R���I�^�J�C�`�U�G�Y�l�~�|�}���|�������D���E�������P�����|�������N�`�d�b�p�c�p�r�O�V�_�~�z�w�t�s�����i�������`�R�Q�U�y���W�_�P�T�`�b�u�w�N�e�u�@�C�c�w�a�I�[Û�{��đ�ٖV�FÄ�z�}Ē�KĚ�Xē�L�_Ó�TĘ�Dā�e�vĜݛŜŞœ�A�D�Gˇ���d�Gʏ�JɐȔ˞�{�O�n�r�K�O�o�d�\�L���O�G�vʁɜ�w�C�jʎ�sȝ��Ο��n�|�p�aȇˎ�W�Rɏ�P�n�W�@�~���L�}Ξ�I�Mʒ�_�[�rʉ�Y�V�{�E�yʚ�v��N�`�A�@�I�N˒�\̔�]̓�A�m�rϊ�gρΛ�Q͘�MϠ�|�U�U͐�u·ϓ͑΁Ϟω�X�sϐ�Nϔ�R���a�rЖ�\�U�m�u�b�dў�cѝ�O�@�hҊ�^ҎҒҕҗ�[�X�JҠ�]�D�M�P�U�x�|�z�u�`ӋӆӇ�J�Iӓӏӑ׌ӘәӖ�hӍӛ�v�M֎�nӠ�G�SӞՓ�A�S�O�L�E�^�b�X�u�{�R�p�V�\�g�a�~�x�t�g�r�E�CԇԟԊԑԜ�\�DԖԒ�QԍԏԎԃԄՊԓԔԌ՟Ԃ�]�_�Z�V�`�a�T�d�N�f�b�OՈ�TՌ�Z�xՎ�u�nՆ՘�lՔ�{�~ՏՁ�rՄ�x�\�Rՙ�e�G�C�o�]�^�@�I�X׋�J�O�V�B�i՛փו�q�x�{�r�u�t�k֔֙ֆ�vև�T�P�S׎�V�Hח�l�d׏ؐؑؓؕؔ؟�t���~؛�|؜؝ؚ�Hُ�A؞�E�v�S�B�N�F�L�J�Q�M�R�O�\ٗ�Z�V�D�U�T�E�Y�W�B�g�c�l�d�xـ�H�p�n�s�r�yهٍَِّ٘�Iٝٛ٠�A�M�w�sڅڎ�O�Sۄ�V�`�Eۋ�]�Q�x�Pۙ�W�U�bۘ�X�f�k�g�|܇܈܉܎ܐ�Dܗ݆ܛ�Z�V�_�S�T�W�F�]�U�p�Y�d�e�I�b�`�^�m�o�v݂݅�x݁�y�z�wݏݗ݋ݔ�\�@ݠݚ�A�H�O�o�q�p߅�|�_�w�^�~�\߀�@�M�h�`�B�t߃ޟ�E�m�x�d�fߊ߉�z�b�����w�]�u�����P�����i�B�y���j�u�����a��Y��������Q�A��C��{�S�O�}���g�n��c�^��k�j耚J�x�u�^��[��^�o�Z��X�`�Q�������X��f�g����F�K��p�U�T��C�B�G���I�D��s�B��e�t�K�~�X��z����b�A�f��|�x��t��P�C�q��P�|�|�@�y��T���n���H�N�i��z���~�n�S�s�h�\�J�R�Z�u�|�H��N�e�^�Q��K�a�d��N�F�\��U�V�I��i�O��|�I�J���@�R��}�D��V�U����k����y��^���\�g�S�M�N��O�R�C�����h��|��j��Z�D�C��O�s��L�T�V�W�Z�]���J�c��e�b�g�h�`���l�[�|�Y�}��y�w�u��b������]�����U�@����H�D�I�R�����A�H��]�����E�U�S�[�`�h�y�rׇ�Z�F�V�\�n�o�v�^�d�f�g�n�t�y�w����������B��D��C��@�A�B�I�H�i�R�a�M�}�W�U�l�j�h�f�w�}����~�D������A�E�L�R�S�Z�`�h�j�w����|�h��q������T�����D���A���G�I�N�H�Q�W�^���t��x�s�}�~�z���R�S�W�Z�Y��g�H�z����x�|�v��w�{�A�~��R������P�G��E�U�T�S�K��_�s�}�\��t�q�~�����E�K�J�t�y�x�W�|�u�~�������Q�|�V�U�c�T�q�^�n�b�q�o�r���\���~�����������������a���N�O�E�H�K�F�T���L���l�w�{�q�v�m�������������B�L�M���Z�X�[�V�k�B�F�u�S�Q�t�f�d�c�����R�����|�z�x�r���v���������[���P�Z�N�]�Z�O���Y�^�o�g�l�i�������X�F�_�Y�Q�W�p�w�������������X�z�����N�S�Z�t�o�w�x���B�R�W�X�Z�e�g�_�f�b�l�r�p�x�}��������"
for i=1 to len(content)
c=mid(content,i,1)

d=instr(t,c)

if d>0 then

temp=temp&mid(s,d,1)

else

temp=temp&c
end if

BigToGb=temp


next

end function


function getpychar(charA)
if len(charA)=0 then exit function
tmp=65536+asc(charA)
'response.write tmp
if(tmp>=45217 and tmp<=45252) then 
getpychar= "A"
elseif(tmp>=45253 and tmp<=45760) then
getpychar= "B"
elseif(tmp>=45761 and tmp<=46317) then
getpychar= "C"
elseif(tmp>=46318 and tmp<=46825) then
getpychar= "D"
elseif(tmp>=46826 and tmp<=47009) then 
getpychar= "E"
elseif(tmp>=47010 and tmp<=47296) then 
getpychar= "F"
elseif(tmp>=47297 and tmp<=47613) then 
getpychar= "G"
elseif(tmp>=47614 and tmp<=48118) then
getpychar= "H"
elseif(tmp>=48119 and tmp<=49061) then
getpychar= "J"
elseif(tmp>=49062 and tmp<=49323) then 
getpychar= "K"
elseif(tmp>=49324 and tmp<=49895) then 
getpychar= "L"
elseif(tmp>=49896 and tmp<=50370) then 
getpychar= "M"
elseif(tmp>=50371 and tmp<=50613) then 
getpychar= "N"
elseif(tmp>=50614 and tmp<=50621) then 
getpychar= "O"
elseif(tmp>=50622 and tmp<=50905) then
getpychar= "P"
elseif(tmp>=50906 and tmp<=51386) then 
getpychar= "Q"
elseif(tmp>=51387 and tmp<=51445) then 
getpychar= "R"
elseif(tmp>=51446 and tmp<=52217) then 
getpychar= "S"
elseif(tmp>=52218 and tmp<=52697) then 
getpychar= "T"
elseif(tmp>=52698 and tmp<=52979) then 
getpychar= "W"
elseif(tmp>=52980 and tmp<=53688) then 
getpychar= "X"
elseif(tmp>=53689 and tmp<=54480) then 
getpychar= "Y"
elseif(tmp>=54481 and tmp<=62289) then
getpychar= "Z"
else '����������ģ��򲻴���
getpychar=charA
end if
end function

%>


