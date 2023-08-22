# Monthly_Report-Excel
Through this project, I wanted to demonstrate what I do on a regular basis in Excel. The main goal was to get through a bunch of invoices, put them in the tables, merge the tables into one and create a simple monthly report. One extra difficulty was to create all those invoices with randomly generated data.
<br>The project contains invoices for food supplies from two companies - Meat and Vegie. From each company, we got ten invoices

### Step zero
To create a random id number I used Excel =RANDBETWEEN() function with 1001 as a lower limit and 4999 as an upper limit. For Unit price column I used the same function but I had to divide it by 100 to get decimal numbers(=RANDBETWEEN(10000,200000)/100). The rest of the columns were simple calculations based on the generated data.

### Step one
Go through the invoices, analyze them, and plan further steps.
![exc1](https://github.com/Bzeorge/Monthly_Report-Excel/assets/74241688/4098d851-6bfb-4d47-abd8-29542b231c79)

### Step two
Formate every invoice as a table and give them specific names containing the date of the invoice which will be useful in later steps. The name of the table contain the company name which issued the invoice and the date of the invoice. It is also very helpful when searching for a specific invoice. 
![exc2](https://github.com/Bzeorge/Monthly_Report-Excel/assets/74241688/87dd330a-20de-46ce-9a54-b0a32bc3b28f)

### Step three
Open the Power Query editor and filter to select only specific company and month - that´s why was so important to name each table by a specific name.
![exc3](https://github.com/Bzeorge/Monthly_Report-Excel/assets/74241688/08660dcd-0123-4c70-abf7-2dcfad409783)

### Step four
Change each column data type so it will be easier to work with them in the next steps. It is especially important to change the data type for currency columns otherwise the numbers would be treated as text. Then I created the Date column by deleting first half of the column name and then formating the column as date data type. In the end, I created a new Company column to easily identify which product belongs to which company. I repeated the process for each company and then I merged both queries together to create one table with all the invoices for the month.
![exc5](https://github.com/Bzeorge/Monthly_Report-Excel/assets/74241688/896751b6-022b-41b3-ba14-bb9b5e6ddb0e)
![exc6](https://github.com/Bzeorge/Monthly_Report-Excel/assets/74241688/c989e850-2a32-4289-bb8f-572cf1067aad)
![exc8](https://github.com/Bzeorge/Monthly_Report-Excel/assets/74241688/7f6de83b-0302-4500-9e1a-8dab9add7157)

### Step five 
Create for each company a few pivot tables summarizing important aspects of the month - a total cost per month,the top 10 most expensive items and a general summary of the products and their amount.
![exc9](https://github.com/Bzeorge/Monthly_Report-Excel/assets/74241688/ceeaea0f-4a90-47d4-895f-57d4b72f1e10)

### Step six
The last step is to create a visual report with the numbers that stakeholders are the most interested in. In this case, it is how much they paid to which company, how much in total, and what product they spent the most money at.
![exc10](https://github.com/Bzeorge/Monthly_Report-Excel/assets/74241688/3b6a9cce-5412-4db5-8a1f-0e835a8c9bb9)

### Conclusion
This report can help stakeholders to get a general overview of the business for the month. Thanks to identifying the top 10 products of the month they can see what items are the most expensive and they can try to find a cheaper supplier for those products.
