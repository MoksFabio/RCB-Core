import sqlite3
import os

def inspect_db(db_path, output_file):
    with open(output_file, 'a', encoding='utf-8') as f:
        f.write(f"--- Inspecting {db_path} ---\n")
        if not os.path.exists(db_path):
            f.write("File not found\n")
            return

        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()
            f.write(f"Tables: {[t[0] for t in tables]}\n")
            
            for table in tables:
                table_name = table[0]
                f.write(f"\nSchema for {table_name}:\n")
                cursor.execute(f"PRAGMA table_info({table_name})")
                columns = cursor.fetchall()
                for col in columns:
                    f.write(str(col) + "\n")
                
                # Show a few rows
                f.write(f"Sample data for {table_name}:\n")
                cursor.execute(f"SELECT * FROM {table_name} LIMIT 3")
                rows = cursor.fetchall()
                for row in rows:
                    f.write(str(row) + "\n")
                    
            conn.close()
        except Exception as e:
            f.write(f"Error: {e}\n")

if __name__ == "__main__":
    if os.path.exists('inspection_result.txt'):
        os.remove('inspection_result.txt')
    # inspect_db('instance/ouvidoria_prazos.db', 'inspection_result.txt')
    inspect_db('instance/logs/log_atividades.db', 'inspection_result.txt')


