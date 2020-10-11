/* Initialize the location of the file */
%LET loc = "P:\MyWork\SAS Questions\office Supplies.xlsx";

/*		1.	Import the “Orders” sheet from the file Office Supplies.xlsx
*/

			PROC IMPORT OUT = ORDERS
			DATAFILE = &loc.					/* &loc. is the defined in Initialization section */
			DBMS = XLSX REPLACE;
			Sheet = Orders;						/* Sheet=  signifies the name of the sheet in excel which needs to be imported*/
			run;

/* 		2.	Make the following changes in the imported dataset as a part of data processing:
*/

/*		2.a. Make the Order ID field as 5 digits (For example – Order ID ‘2’ should be shown as ‘00002’
*/
			DATA ORDERS1;
			SET ORDERS;
			Order_ID1 = put(Order_ID,z5.); 			/* z5. is the main functionality which sets the number of digits as "5"  */
			DROP Order_ID;
			RENAME Order_ID1=Order_ID;
			RUN;


/*		2.b. Product Base Margin should be a numeric field. Please remove “na” convert the field to numeric (na should just be blank now)
*/
			DATA ORDERS2;
			SET ORDERS1;
			IF Product_Base_Margin = 'na' then Product_Base_Margin =""; 
			Product_Base_Margin1 = input(Product_Base_Margin,4.); 				/* Input function converts String to Numerics */
			drop Product_Base_Margin;
			rename Product_Base_Margin1 = Product_Base_Margin;
			RUN;

/*		3.	Import the ‘States’ sheet from the same file which has the mapping for State field in Orders sheet. 
*/

			PROC IMPORT OUT = States
			DATAFILE = &loc.				/* &loc. is the defined in Initialization section */
			DBMS = XLSX REPLACE;
			Sheet = States;       			 /* Sheet=  signifies the name of the sheet in excel which needs to be imported*/
			run;


/*		4.	Merge the State dataset to the Orders dataset to get the State_Name in your order dataset
*/

/*		4.a. Sorting the Orders dataset before merging with State dataset to get the StateName
			In SAS, before any Merge Operation, Sorting of the individual dataset is necessary
*/

			PROC SORT DATA = ORDERS2 OUT= ORDERS2_SORTED;
			BY
			State Row_ID Order_ID Order_Date Order_Priority Order_Quantity Sales Discount Ship_Mode Profit Unit_Price Shipping_Cost Customer_Name
			Product_Sub_Category Product_Name Product_Container Product_Base_Margin Ship_Date;
			RUN;


/*		4.b. Sorting the State dataset before merging with Orders dataset to get the StateName
			In SAS, before any Merge Operation, Sorting of the individual dataset is necessary
*/

			PROC SORT DATA = States OUT= States_SORTED;
			BY State;
			RUN;

/*		4.c. Merging the two datasets: States dataset and Orders dataset to get the StateName			
*/

			DATA ORDERS2_STATES;
			MERGE ORDERS2_SORTED(in=A) States_SORTED (in=B);
			by State;
			IF A OR B;  						/* We have assumed here we have to consider all the cases wherein there is a State ID */
			RUN;								/*	hence we have used A or B	*/

/*		5.	Calculate the 
			a. 	Total Profit,		
			b. 	Number of Orders and
			c. 	Average Profit
				by State_Name and Ship Mode. 
				Create a dataset called State_Ship containing this information. 
*/

/*		5.a Total Profit
*/
			PROC SUMMARY DATA = ORDERS2_STATES NWAY MISSING  ;
			var Profit;
			class State_Name Ship_mode;
			OUTPUT OUT = Total_Profit (DROP = _TYPE_ _FREQ_)SUM= ;
			run;

/*		5.b Number of Orders	
*/
			PROC SQL;
			Create Table Order_no as
			Select State_Name,Ship_mode,count(distinct order_id) AS No_of_Orders from ORDERS2_STATES group by 1,2;
			QUIT;

/*		5.c Average Profit		
			To calculate Average Profit, I assumed that Profit needs to be calculated based on a Order.
			Hence, Average Profit is Profit per individual order

			It would be better to merge the above two datasets to arrive at the Average Profit 
			as Average Profit = Total Profit/No_of_Orders
*/
			PROC SQL;
			CREATE TABLE PROFIT_ORDER AS
			SELECT A.State_Name, A.Ship_Mode, A.Profit, B.No_of_Orders from Total_Profit as A INNER JOIN
			Order_no as B ON A.State_Name=B.State_Name AND A.Ship_Mode= B.Ship_Mode;
			QUIT;


           /* Calculate Average Profit */
			DATA State_Ship;
			SET PROFIT_ORDER;
			Average_Profit = Profit/No_of_orders;
			run;


/*        6.For each Product Sub Category, remove the record which has the highest Sales. 
			Create a dataset called Orders_sub for this.
 
			(For example – for product sub category Appliances, delete the record where 
			Sales is 16002.29 which is the highest for that sub category).
*/

/*		   6.a. Sorting the raw dataset ORDERS2_STATES based on sales and Product_Sub_Category from Smallest and Largest
*/
			PROC SORT DATA = ORDERS2_STATES out = Orders_sub1; by Product_Sub_Category Sales; run;

/*		   6.b. Remove the first and last occurence of each Product_Sub_Category as Orders_sub1 is already sorted in that way
*/

			DATA Orders_sub;
			SET Orders_sub1;
			by Product_Sub_Category Sales;
			IF first.Product_Sub_Category then flag=1;
			IF last.Product_Sub_Category then flag=1;
			run;

/*		   6.c. Delete the cases wherein Flag=1 as this would eliminate those cases wherein Product Sales Category has Sales value as minimum and
			maximum
*/

			DATA Orders_sub;
			SET Orders_sub;
			If Flag=1 then delete;
			run;


/* 			7.	In the original dataset Orders, replace all the occurrences of the word ‘Fellowes’ with ‘Follows’ in the field Product Name
*/

/*			To Cater this problem, we will be using the SAS function called TRANWRD to replace "Fellowes" with "Follows"
*/

			DATA ORDERS2_STATES;
			SET ORDERS2_STATES;
			Product_Name1 = TRANWRD(Product_Name,"Fellowes","Follows");
			DROP Product_Name;
			RENAME Product_Name1=Product_Name;
			RUN;

/* 			8. Filter for customers whose names contain the word ‘Jack’ and calculate the total ‘Order Quantity’ for these customers. 
			   Store the information in a dataset called jack_qty
*/

/*			We are assuming here, if the Customer_Name has a word "Jack" anywhere, either in the first name or in the last name, we 
			will consider them to calculate "Order Quantity"
*/

			PROC SQL;
			CREATE TABLE jack_qty AS
			SELECT Customer_Name, SUM(Order_Quantity) as Order_Quantity FROM ORDERS2_STATES where Customer_Name like "%Jack%" group by 1 ;
			QUIT;



