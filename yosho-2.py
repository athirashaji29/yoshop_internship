#!/usr/bin/env python
# coding: utf-8

# In[61]:


order_data


# In[62]:


order_data.columns


# In[47]:


df = pd.DataFrame(order_data)

# Display specific column values
selected_column = 'Payment Date and Time Stamp'  # Specify the column you want to display
column_values = df[selected_column]

print(column_values)


# In[2]:


get_ipython().system('pip install reportlab')


# In[63]:


review_data


# In[59]:


review_data.columns


# In[3]:


from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


# In[60]:


import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import datetime

# Load the review data and order data into separate DataFrames
review_data = pd.read_csv('review_dataset.csv')
order_data = pd.read_csv('ordersDataset.csv')

# Function to generate analysis report in PDF and Excel format based on user's choice
def generate_analysis_report(choice):
    if choice == 1:
        # Analysis of Reviews given by Customers
        # Perform the analysis and generate the chart
        review_counts = review_data['stars'].value_counts()
        plt.figure(figsize=(10, 6))
        sns.barplot(x=review_counts.index, y=review_counts.values)
        plt.xlabel('Review')
        plt.ylabel('Count')
        plt.title('Analysis of Reviews given by Customers')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig('reviews_analysis.png')

        # Generate the analysis report in PDF and Excel format
        generate_report('Reviews given by Customers', review_counts, 'reviews_analysis.pdf', 'reviews_analysis.xlsx')

    elif choice == 2:
        # Analysis of Different Payment Methods used by the Customers
        # Perform the analysis and generate the chart
        payment_counts = order_data['Payment Method'].value_counts()
        plt.figure(figsize=(10, 6))
        sns.barplot(x=payment_counts.index, y=payment_counts.values)
        plt.xlabel('Payment Method')
        plt.ylabel('Count')
        plt.title('Analysis of Different Payment Methods used by the Customers')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig('payment_methods_analysis.png')

        # Generate the analysis report in PDF and Excel format
        generate_report('Different Payment Methods used by Customers', payment_counts, 'payment_methods_analysis.pdf', 'payment_methods_analysis.xlsx')

    elif choice == 3:
        # Analysis of Top Consumer States of India
        # Perform the analysis and generate the chart
        top_states = order_data['Shipping State'].value_counts().head(10)
        plt.figure(figsize=(10, 6))
        sns.barplot(x=top_states.index, y=top_states.values)
        plt.xlabel('State')
        plt.ylabel('Count')
        plt.title('Analysis of Top Consumer States of India')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig('top_states_analysis.png')

        # Generate the analysis report in PDF and Excel format
        generate_report('Top Consumer States of India', top_states, 'top_states_analysis.pdf', 'top_states_analysis.xlsx')

    elif choice == 4:
        # Analysis of Top Consumer Cities of India
        # Perform the analysis and generate the chart
        top_cities = order_data['Shipping City'].value_counts().head(10)
        plt.figure(figsize=(10, 6))
        sns.barplot(x=top_cities.index, y=top_cities.values)
        plt.xlabel('City')
        plt.ylabel('Count')
        plt.title('Analysis of Top Consumer Cities of India')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig('top_cities_analysis.png')

        # Generate the analysis report in PDF and Excel format
        generate_report('Top Consumer Cities of India', top_cities, 'top_cities_analysis.pdf', 'top_cities_analysis.xlsx')

    elif choice == 5:
        # Analysis of Top Selling Product Categories
        # Perform the analysis and generate the chart
        top_categories = review_data['category'].value_counts().head(10)
        plt.figure(figsize=(10, 6))
        sns.barplot(x=top_categories.index, y=top_categories.values)
        plt.xlabel('Product Category')
        plt.ylabel('Count')
        plt.title('Analysis of Top Selling Product Categories')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig('top_categories_analysis.png')

        # Generate the analysis report in PDF and Excel format
        generate_report('Top Selling Product Categories', top_categories, 'top_categories_analysis.pdf', 'top_categories_analysis.xlsx')
    elif choice == 6:
        # Analysis of Reviews for All Product Categories
        # Perform the analysis and generate the chart
        review_counts = review_data.groupby('category')['stars'].value_counts().unstack(fill_value=0)
        review_counts.plot(kind='bar', figsize=(10, 6))
        plt.xlabel('Product Category')
        plt.ylabel('Count')
        plt.title('Analysis of Reviews for All Product Categories')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig('reviews_all_categories.png')

        # Generate the analysis report in PDF and Excel format
        generate_report('Reviews for All Product Categories', review_counts, 'reviews_all_categories.pdf', 'reviews_all_categories.xlsx')
    elif choice == 7:
        # Analysis of Number of Orders Per Month Per Year
        # Perform the analysis and generate the chart
        order_data['order_date'] = pd.to_datetime(order_data['Order Date and Time Stamp'])
        order_data['year_month'] = order_data['order_date'].dt.to_period('M')
        order_counts = order_data['year_month'].value_counts().sort_index()
        plt.figure(figsize=(10, 6))
        order_counts.plot(kind='bar')
        plt.xlabel('Year-Month')
        plt.ylabel('Number of Orders')
        plt.title('Analysis of Number of Orders Per Month Per Year')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig('orders_per_month.png')

        # Generate the analysis report in PDF and Excel format
        generate_report('Number of Orders Per Month Per Year', order_counts, 'orders_per_month.pdf', 'orders_per_month.xlsx')
    elif choice == 8:
        # Analysis of Reviews for Number of Orders Per Month Per Year
        # Perform the analysis and generate the chart
        order_data['order_date'] = pd.to_datetime(order_data['Order Date and Time Stamp'])
        order_data['year_month'] = order_data['order_date'].dt.to_period('M')
        review_order_counts = order_data.groupby('year_month')['Order #'].count()
        plt.figure(figsize=(10, 6))
        review_order_counts.plot(kind='bar')
        plt.xlabel('Year-Month')
        plt.ylabel('Number of Orders')
        plt.title('Analysis of Reviews for Number of Orders Per Month Per Year')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig('reviews_per_month.png')

        # Generate the analysis report in PDF and Excel format
        generate_report('Reviews for Number of Orders Per Month Per Year', review_order_counts, 'reviews_per_month.pdf', 'reviews_per_month.xlsx')
    elif choice == 9:
        # Analysis of Number of Orders Across Parts of a Day
        # Perform the analysis and generate the chart
        order_data['order_time'] = pd.to_datetime(order_data['Order Date and Time Stamp']).dt.time
        parts_of_day = []
        for time in order_data['order_time']:
            if time < datetime.time(6, 0, 0):
                parts_of_day.append('Night')
            elif time < datetime.time(12, 0, 0):
                parts_of_day.append('Morning')
            elif time < datetime.time(18, 0, 0):
                parts_of_day.append('Afternoon')
            else:
                parts_of_day.append('Evening')
        order_data['part_of_day'] = parts_of_day

        order_counts_day = order_data['part_of_day'].value_counts()
        plt.figure(figsize=(10, 6))
        sns.barplot(x=order_counts_day.index, y=order_counts_day.values)
        plt.xlabel('Part of Day')
        plt.ylabel('Number of Orders')
        plt.title('Analysis of Number of Orders Across Parts of a Day')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig('orders_per_part_of_day.png')

        # Generate the analysis report in PDF and Excel format
        generate_report('Number of Orders Across Parts of a Day', order_counts_day, 'orders_per_part_of_day.pdf', 'orders_per_part_of_day.xlsx')
    elif choice == 10:
        # Full Report
        # Perform the analysis and generate the charts and reports for all available choices
        generate_analysis_report(1)  # Reviews given by Customers
        generate_analysis_report(2)  # Different Payment Methods used by the Customers
        generate_analysis_report(3)  # Top Consumer States of India
        generate_analysis_report(4)  # Top Consumer Cities of India
        generate_analysis_report(5)  # Top Selling Product Categories
        generate_analysis_report(6)  # Reviews for All Product Categories
        generate_analysis_report(7)  # Number of Orders Per Month Per Year
        generate_analysis_report(8)  # Reviews for Number of Orders Per Month Per Year
        generate_analysis_report(9)  # Number of Orders Across Parts of a Day


   
# Function to generate the analysis report in PDF and Excel format
def generate_report(title, data, pdf_filename, excel_filename):
    # Generate PDF report
    pdf = canvas.Canvas(pdf_filename, pagesize=letter)
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(30, 750, title)
    pdf.setFont("Helvetica", 12)
    y = 700
    for index, (label, value) in enumerate(data.items()):
        pdf.drawString(30, y, f'{label}: {value}')
        y -= 20
        if index % 30 == 0 and index != 0:
            pdf.showPage()
            y = 750
    pdf.save()

    # Generate Excel report
    df = pd.DataFrame(data, columns=['Label', 'Value'])
    df.to_excel(excel_filename, index=False)

# Main function
def main():
    # User's choice
    choice = int(input('Enter the number to see the analysis of your choice: '))

    try:
        generate_analysis_report(choice)
        print('Analysis report generated successfully.')
    except Exception as e:
        print(f'An error occurred: {str(e)}')

if __name__ == '__main__':
    main()


# In[ ]:





# In[ ]:




