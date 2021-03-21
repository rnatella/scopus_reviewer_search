# Dependencies

```
python3 -m venv env
source env/bin/activate
pip install -r requirements.txt
```

# Reviewer search by keywords (AND/OR separated)

```
python3 scopus_search.py -k "fault injection"
```

# Reviewer search by references (JSON format generated from anystyle)

```
sudo gem install anystyle-cli
anystyle find file.pdf > references.json

python3 scopus_search.py -j references.json
```

# Reviewer search by references (TXT format, plain titles)

python3 scopus_search.py -t references.txt
