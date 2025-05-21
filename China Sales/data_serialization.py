import psycopg2
import json

def get_data():
    # 连接 PostgreSQL 数据库
    conn = psycopg2.connect(
        host="127.0.0.1",
        database="CS",
        user="postgres",
        password="Radioflyer1",
        port="5432"
    )
    print("Database connected.")
    cur = conn.cursor()

    # 查询数据
    cur.execute("""
        SELECT 
            "rma#" as rma,
            to_char(contact_date,'YYYY-MM-DD') as contact_date,
            to_char(purchase_date,'YYYY-MM-DD') as purchase_date,
            contact_id,
            source_of_purchase,
            "item#" as item,
            defect_unit,
            original_address,
            defect_description,
            action_to_be_taken,
            parts_no,
            "tracking#" as tracking,
            courier_,
            complaint_category_class_i,
            complaint_category_class_ii,
            factory,
            name,
            number,
            address
        FROM 
            after_sales_records
    """)
    data = cur.fetchall()
    print([dict(zip([column[0] for column in cur.description], row)) for row in data])
    # 将数据转换为 JSON 格式
    json_data = json.dumps([dict(zip([column[0] for column in cur.description], row)) for row in data],ensure_ascii=False)
    
    # 关闭数据库连接
    cur.close()
    conn.close()
    return json_data

