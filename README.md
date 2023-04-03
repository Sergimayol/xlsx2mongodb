# .xlsx to MongoDB

This is a simple script to transfer data from an Excel file to a MongoDB database.

## Usage

Inside the `src/main.py` file, you can change the following variables:

```python
DB_NAME = "" # Database name
URI = "mongodb://localhost:27017" # MongoDB URI
DIR_PATH = r"" # Path to the directory where the .xlsx file/s is/are located
```

Then, you can run the script with the following command:

```bash
cd src
python main.py
```

## Requirements

```bash
pip install -r requirements.txt
```

## License

[MIT](./LICENSE)
