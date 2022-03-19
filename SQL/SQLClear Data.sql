
Delete AccountMaster Where Code <>'000000' Or Code <>'*00001' And [Group] Not IN ('*99996','*99997','*99998','*99999')
Delete from BookDNParent
Select * from BookingRouteMaster
Delete From BookMaster
Delete from BookingRouteMaster
Delete from BookOOParent
Delete from BookPOParent
Delete from BookRVParent
Delete from JobworkBVParent
Delete from MaterialIOParent
Delete from MaterialMVParent
Delete from MaterialSVParent
Delete from OutsourceItemPOParent
Delete from PackingSlipParent
Delete from PaperDNParent
Delete from PaperMVParent
Delete from PaperPOParent
Delete from PrintPVParent
Delete from TatRVParent
----Delete From OutsourceItemMaster
----Delete From PaperMaster

Select *From UserMaster
Delete From UserMaster Where Code Not IN ('000001','000005') And [Level]=1
Delete From UserAction
Delete From TeamMemberMaster

Select *From GeneralMaster Where Type IN ('5') 
Delete From GeneralMaster Where Type IN ('5') 

--------------------------ElementMaster Last Code *00045
Select *from ElementMaster
-------------------------------------------------Account Master-000000------------------------------------------------
Update AccountMaster Set 
Code='000000',
Name= (Select Name from CompanyMaster),
PrintName= (Select Name from CompanyMaster),
Alias='1001',
Address1= (Select Address1 from CompanyMaster),
Address2= (Select Address2 from CompanyMaster),
Address3= (Select Address3 from CompanyMaster),
Address4= (Select Address4 from CompanyMaster),
Phone= (Select Phone from CompanyMaster),
Mobile= (Select Mobile from CompanyMaster),
TIN= (Select GSTIN from CompanyMaster),
eMail= (Select eMail from CompanyMaster),
RoundOffQty= 1,
CreatedBy='000001',
CreatedOn= (Select CreatedOn from AccountMaster Where Code='000000'),
ModifiedBy= (Select ModifiedBy from AccountMaster Where Code='000000'),
ModifiedOn= (Select ModifiedOn from AccountMaster Where Code='000000'),
Recordstatus= (Select Recordstatus from AccountMaster Where Code='000000'),
Printstatus= (Select Printstatus from AccountMaster Where Code='000000')
Where Code='000000'

-----Update AccountMaster Set Name=Replace(Name,'Kaveri - ','KOPL - ')

-------------------------------------------------Account Master-*00001------------------------------------------------
Insert Into AccountMaster VALUES 
('*00001',
'Rate Master',
'Rate Master',
'1002',
'001090',
(Select Address1 from CompanyMaster),
(Select Address2 from CompanyMaster),
(Select Address3 from CompanyMaster),
(Select Address4 from CompanyMaster),
(Select Phone from CompanyMaster),
(Select Mobile from CompanyMaster),
(Select GSTIN from CompanyMaster),
(Select eMail from CompanyMaster),
 1,
'000001',
(Select CreatedOn from AccountMaster Where Code='000000'),
(Select ModifiedBy from AccountMaster Where Code='000000'),
(Select ModifiedOn from AccountMaster Where Code='000000'),
(Select Recordstatus from AccountMaster Where Code='000000'),
(Select Printstatus from AccountMaster Where Code='000000'));


-------------------------------------------------Account Master-*00002------------------------------------------------

Insert Into AccountMaster VALUES 
('*00002',
'Main Godown',
'Main Godown',
'1002',
'*99999',
(Select Address1 from CompanyMaster),
(Select Address2 from CompanyMaster),
(Select Address3 from CompanyMaster),
(Select Address4 from CompanyMaster),
(Select Phone from CompanyMaster),
(Select Mobile from CompanyMaster),
(Select GSTIN from CompanyMaster),
(Select eMail from CompanyMaster),
 1,
'000001',
(Select CreatedOn from AccountMaster Where Code='000000'),
(Select ModifiedBy from AccountMaster Where Code='000000'),
(Select ModifiedOn from AccountMaster Where Code='000000'),
(Select Recordstatus from AccountMaster Where Code='000000'),
(Select Printstatus from AccountMaster Where Code='000000'));

-------------------------------------------------Account Master-*00003------------------------------------------------
Insert Into AccountMaster VALUES 
('*00003',
'Direct',
'Direct',
'1003',
'*99998',
(Select Address1 from CompanyMaster),
(Select Address2 from CompanyMaster),
(Select Address3 from CompanyMaster),
(Select Address4 from CompanyMaster),
(Select Phone from CompanyMaster),
(Select Mobile from CompanyMaster),
(Select GSTIN from CompanyMaster),
(Select eMail from CompanyMaster),
 1,
'000001',
(Select CreatedOn from AccountMaster Where Code='000000'),
(Select ModifiedBy from AccountMaster Where Code='000000'),
(Select ModifiedOn from AccountMaster Where Code='000000'),
(Select Recordstatus from AccountMaster Where Code='000000'),
(Select Printstatus from AccountMaster Where Code='000000'));

-------------------------------------------------Account Master-*00004------------------------------------------------
Insert Into AccountMaster VALUES 
('*00004',
'Packer',
'Packer',
'1004',
'*99997',
(Select Address1 from CompanyMaster),
(Select Address2 from CompanyMaster),
(Select Address3 from CompanyMaster),
(Select Address4 from CompanyMaster),
(Select Phone from CompanyMaster),
(Select Mobile from CompanyMaster),
(Select GSTIN from CompanyMaster),
(Select eMail from CompanyMaster),
 1,
'000001',
(Select CreatedOn from AccountMaster Where Code='000000'),
(Select ModifiedBy from AccountMaster Where Code='000000'),
(Select ModifiedOn from AccountMaster Where Code='000000'),
(Select Recordstatus from AccountMaster Where Code='000000'),
(Select Printstatus from AccountMaster Where Code='000000'));

-------------------------------------------------Account Master-*00005------------------------------------------------
Insert Into AccountMaster VALUES 
('*00005',
'Self Transport',
'Self Transport',
'1005',
'*99996',
(Select Address1 from CompanyMaster),
(Select Address2 from CompanyMaster),
(Select Address3 from CompanyMaster),
(Select Address4 from CompanyMaster),
(Select Phone from CompanyMaster),
(Select Mobile from CompanyMaster),
(Select GSTIN from CompanyMaster),
(Select eMail from CompanyMaster),
 1,
'000001',
(Select CreatedOn from AccountMaster Where Code='000000'),
(Select ModifiedBy from AccountMaster Where Code='000000'),
(Select ModifiedOn from AccountMaster Where Code='000000'),
(Select Recordstatus from AccountMaster Where Code='000000'),
(Select Printstatus from AccountMaster Where Code='000000'));

-------------------------------------------------General Master-******------------------------------------------------
Insert Into GeneralMaster VALUES ('*10001','General','General',5,0,'000001',(Select CreatedOn from AccountMaster Where Code='000000'),(Select ModifiedBy from AccountMaster Where Code='000000'),(Select ModifiedOn from AccountMaster Where Code='000000'),(Select Recordstatus from AccountMaster Where Code='000000'),(Select Printstatus from AccountMaster Where Code='000000'));
-------------------------------------------------General Master-******------------------------------------------------
Insert Into GeneralMaster VALUES ('*10002','Account Group','Account Group',12,0,'000002',(Select CreatedOn from AccountMaster Where Code='000000'),(Select ModifiedBy from AccountMaster Where Code='000000'),(Select ModifiedOn from AccountMaster Where Code='000000'),(Select Recordstatus from AccountMaster Where Code='000000'),(Select Printstatus from AccountMaster Where Code='000000'));
-------------------------------------------------General Master-******------------------------------------------------
Insert Into GeneralMaster VALUES ('*10003','Debtors','Debtors',12,0,'000003',(Select CreatedOn from AccountMaster Where Code='000000'),(Select ModifiedBy from AccountMaster Where Code='000000'),(Select ModifiedOn from AccountMaster Where Code='000000'),(Select Recordstatus from AccountMaster Where Code='000000'),(Select Printstatus from AccountMaster Where Code='000000'));
-------------------------------------------------General Master-******------------------------------------------------
Insert Into GeneralMaster VALUES ('*10004','Creditor','Creditor',12,0,'000004',(Select CreatedOn from AccountMaster Where Code='000000'),(Select ModifiedBy from AccountMaster Where Code='000000'),(Select ModifiedOn from AccountMaster Where Code='000000'),(Select Recordstatus from AccountMaster Where Code='000000'),(Select Printstatus from AccountMaster Where Code='000000'));
-------------------------------------------------General Master-******------------------------------------------------
Insert Into GeneralMaster VALUES ('*10005','Binders','Binder',12,0,'000005',(Select CreatedOn from AccountMaster Where Code='000000'),(Select ModifiedBy from AccountMaster Where Code='000000'),(Select ModifiedOn from AccountMaster Where Code='000000'),(Select Recordstatus from AccountMaster Where Code='000000'),(Select Printstatus from AccountMaster Where Code='000000'));
-------------------------------------------------General Master-******------------------------------------------------
Insert Into GeneralMaster VALUES ('*10006','Printers','Printers',12,0,'000006',(Select CreatedOn from AccountMaster Where Code='000000'),(Select ModifiedBy from AccountMaster Where Code='000000'),(Select ModifiedOn from AccountMaster Where Code='000000'),(Select Recordstatus from AccountMaster Where Code='000000'),(Select Printstatus from AccountMaster Where Code='000000'));

---Delete From GeneralMaster where code='*10005'

--------------------------------------------------------VchSeriesMaster
--VchSeriesMaster to be update

Update CompanyMaster Set FinancialYearFrom='2021-04-01 00:00:00.000' ,FinancialYearTo='2022-03-31 00:00:00.000',CreatedFrom=''