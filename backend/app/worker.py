import time

from .settings import get_setting
from .db import db, init_db, is_postgres
from .indexer import index_document


def main():
    init_db()
    while True:
        with db() as conn:
            row = conn.execute("SELECT id, document_id FROM index_jobs WHERE status = 'pending' ORDER BY id LIMIT 1").fetchone()
            if row:
                job_id = row["id"]
                document_id = row["document_id"]
                conn.execute("UPDATE index_jobs SET status = 'processing', attempts = attempts + 1 WHERE id = %s" if is_postgres() else "UPDATE index_jobs SET status = 'processing', attempts = attempts + 1 WHERE id = ?", (job_id,))
            else:
                job_id = None
        if not job_id:
            time.sleep(float(get_setting("index_poll_seconds")))
            continue
        try:
            index_document(document_id)
            with db() as conn:
                conn.execute("UPDATE index_jobs SET status = 'done' WHERE id = %s" if is_postgres() else "UPDATE index_jobs SET status = 'done' WHERE id = ?", (job_id,))
        except Exception as exc:
            with db() as conn:
                conn.execute("UPDATE index_jobs SET status = 'failed', error = %s WHERE id = %s" if is_postgres() else "UPDATE index_jobs SET status = 'failed', error = ? WHERE id = ?", (str(exc), job_id))
                conn.execute("UPDATE documents SET index_status = 'failed', index_error = %s WHERE id = %s" if is_postgres() else "UPDATE documents SET index_status = 'failed', index_error = ? WHERE id = ?", (str(exc), document_id))


if __name__ == "__main__":
    main()
