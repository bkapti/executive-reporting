import pymysql
import openpyxl
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from credentials import (
    db,
    user_name,
    user_password,
    remote_port,
    sender_email,
    receiver_email,
    password
)
from email_sender import send_email


def run_query(query):
    try:
        conn = pymysql.connect(
            user=user_name,
            password=user_password,
            host=db,
            db="dbs",
            port=remote_port,
        )
        try:
            data_frame = pd.read_sql_query(query, conn)
            conn.close()
            return data_frame
        except (
                pymysql.err.InternalError,
                pymysql.err.DatabaseError,
                pd.io.sql.DatabaseError,
        ) as e:
            print(
                f'Error occurred while executing a query "{query}", with: {e}'
            )
            conn.close()
            raise

    except pymysql.err.ProgrammingError as except_detail:
        print(f"pymysql.err.ProgrammingError: «{except_detail}»")
        raise pymysql.err.ProgrammingError(except_detail)


def enforce_date():
    # Since, date is in the past, we're hard coding it here.
    todays_date = datetime(2018, 9, 23)
    last_years_date = todays_date - relativedelta(years=1)
    todays_date = todays_date.strftime("%Y-%m-%d")
    last_years_date = last_years_date.strftime("%Y-%m-%d")
    return todays_date, last_years_date


def get_query():
    query = f""" 
            select 
            year(date) as years, 
            week(date,1) as weeks, 
            round(sum(cost)+ sum(datacost)+sum(revenue)) as gross_revenue, 
            round((sum(cost)+ sum(datacost)+sum(revenue))/count(distinct date)) as daily_average_gross_revenue,
            round(sum(revenue)) as net_revenue,
            round(sum(revenue)/count(distinct date)) as daily_average_net_revenue,
            round(sum(revenue)/ sum(cost)+ sum(datacost)+sum(revenue)) as margin,
            round(((sum(cost)+ sum(datacost)+sum(revenue)-
                lead(sum(cost)+ sum(datacost)+sum(revenue)) over (order by year(date), week(date,1) desc))
                    / lead(sum(cost)+ sum(datacost)+sum(revenue)) over (order by year(date), week(date,1) desc))*100) 
                        as period_over_period_growth_gross_revenue,
            if(year(date)=2017,NULL,round(((sum(cost)+ sum(datacost)+sum(revenue)-
                lead(sum(cost)+ sum(datacost)+sum(revenue)) over (order by week(date,1), year(date) desc))
                    / lead(sum(cost)+ sum(datacost)+sum(revenue)) over (order by week(date,1), year(date) desc))*100)) 
                        as year_over_year_growth_gross_revenue,
            round(((sum(revenue)-lead(sum(revenue)) over (order by year(date), week(date,1) desc))
                / lead(sum(revenue)) over (order by year(date), week(date,1) desc))*100)
                    as period_over_period_growth_net_revenue,
            if(year(date)=2017,NULL,round(((sum(revenue)-lead(sum(revenue)) over (order by week(date,1), year(date) desc))
                / lead(sum(revenue)) over (order by week(date,1), year(date) desc))*100)) 
                    as year_over_year_growth_net_revenue
            from fact_date_customer_campaign f join dim_customer d on f.customer_id = d.customer_id 
            where 
            segment = "Segment A"  
            and (
                date between '2018-01-01' and '{todays_date}' 
                    or
                date between '2017-01-01' and '{last_years_date}' 
                )
            group by years, weeks
            having weeks > 0
            order by years desc, weeks desc;
        """
    return query


def beautify_excel(filename):
    wb = openpyxl.load_workbook(filename=filename)
    worksheet = wb.active
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column].width = adjusted_width
    wb.save(f"{todays_date}.xlsx")


if __name__ == "__main__":

    todays_date, last_years_date = enforce_date()
    query = get_query()
    df = run_query(query)
    df.to_excel(f"{todays_date}.xlsx", index=False)
    beautify_excel(f"{todays_date}.xlsx")

    # Set up subject and body
    subject = "Daily report - DSP Product"
    body = "This report is generated automatically by Data Engineering Team."

    send_email(subject=subject,
               body=body,
               sender_email=sender_email,
               receiver_email=receiver_email,
               password=password,
               file_name=f"{todays_date}.xlsx"
               )

