Outlook���ε{�������G

���{���i�H�̷Ӭd�߱���A�˯��t�a�ݻP���ݦ���X���Ҧ���Ƨ����H��A�åHxlsx�ɲ��X�C
�˯��y�k�}��s���Fitrer.txt�A��d�߻y����DASL�A�B���ɮ׽s�X��UTF-8�C

�H�U�������y�k�����G

1. urn:schemas:httpmail:subject�G��ܥD���C
2. urn:schemas:httpmail:textdescription�G��ܤ���C
3. urn:schemas:httpmail:datereceived�G��ܫH�󱵦�����C
4. urn:schemas:httpmail:hasattachment�G��ܪ��ɼƶq�C
5. �h������i�H�ϥ�OR�PAND�A���C�@�ӫ��O�����Ρu()�v�]�СA�_�h�{���i�ण�|�z�|�C
6. �w���ҽk�d��(Like)�᭱�a����r�n�Ρu'% �v�P�u%'�v�]�СA��L�h�Ρu'�v�P�u'�v�]�ЧY�i�C
7. �u�i�j�v�����Ѯج[�A�Y�ؿ�d�򤺬����ѡA���d�N���i�h�h�ؿ�(������DASL�W�h�A�����{�����w)�C
8. �i����������Y�ơC
9. �H�U���d�߻y�k�d�ҡG
@SQL=
(
	("urn:schemas:httpmail:subject" Like \'%Receive_Purchase_Orders%\')�i�D���j
	OR ("urn:schemas:httpmail:textdescription" Like \'%�u��إߵ��G%\')�i����j
) 
AND ("urn:schemas:httpmail:datereceived" > \'2022/04/01\')�i��������j
AND ("urn:schemas:httpmail:datereceived" < \'2022/06/01\')
AND ("urn:schemas:httpmail:hasattachment" = 0)�i���ɼƶq�j