delete from Sales_Transactions where Transaction_Number in
(select Transaction_Number from [192.168.1.206].POS_LOCAL.DBO.Sales_Transactions where Upload_Status ='00')

delete from Sales_Transaction_details where Transaction_Number in
(select Transaction_Number from [192.168.1.206].POS_LOCAL.DBO.Sales_Transactions where Upload_Status ='00')

delete from paid where Transaction_Number in
(select Transaction_Number from [192.168.1.206].POS_LOCAL.DBO.Sales_Transactions where Upload_Status ='00')
