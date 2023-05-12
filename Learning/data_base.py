import psycopg2

con = psycopg2.connect(
database="lottery",
user="lottery",
password="lottery",
host="postgres1.dev.int.nl-dev.ru",
port="5432"
)
print("Database opened successfully")

cur = con.cursor()
cur.execute("""
        SELECT number FROM number_handler.number_ticket, scheduler.draw 
    where 
        draw.id = number_ticket.draw_id
    
        and is_sold = 'true' 
        and is_winning = 'true' 
        and payment_status = 'NOT_PAID'  
        and date_when_unpaid > current_timestamp
        and pay_transaction_id is null
        and hashcode_password = '6R6izWL9kBs='
        and number like '%'
    ORDER BY random()
    LIMIT 5
    """)

print(cur.fetchall())

rows = cur.fetchall()
for row in rows:
   print("NUMBER =", row[0], "\n")

con.close()
