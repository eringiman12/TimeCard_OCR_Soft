import MySQLdb

# DB関係のクラス
class DB_Operation():
    def __init__(self):
        # 接続する
        self.con = MySQLdb.connect(
            user='root',
            passwd='',
            host='localhost',
            db='ocr_sys_db')

    def Select_value(self, Get_sql):
        # カーソルを取得する
        cur = self.con.cursor()
        # SQL（データベースを操作するコマンド）を実行する
        # userテーブルから、HostとUser列を取り出す
        sql = Get_sql
        cur.execute(sql)

        # 実行結果を取得する
        rows = cur.fetchall()

        # カーソルを閉じる
        cur.close
        # 接続を閉じる
        self.con.close

        return rows
        