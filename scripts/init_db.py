import os
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from schema_loader import load_schema
from db import seed_schema

schema_path = os.getenv('SCHEMA_PATH', str(Path(__file__).resolve().parents[1] / 'data' / 'schema.xlsx'))
schema = load_schema(schema_path)
seed_schema(schema)
print('Database initialized and schema seeded.')
