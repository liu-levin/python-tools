# python-tools
Scripts for handling some daily needs

##### use requirements.txt
```
pip install -r requirement.txt
pip freeze >requirements.txt

pip freeze >requirements.txt
pipreqs ./
pipreqs ./ --force
```


#### crawler
- douban_ren.py
  1. Help a friend find rent infos  in douban group
  2. Use xlsxwriter to write xlsx and requests to fetch html content
  3. Use re to `String Pattern Matching`
  4. Crawler used cookie from chrome，and the automated fetching did not complete
  5. To be optimized

      

#### document
- convert.py
  1. Help a friend  convert  xlsx to docx, the format conversion is poor.
  2. Usew xlrd to read xlsx. xlrd-2.0.1 not support xlsx, please use 1.2.0
  ```
    pip uninstall xlrd
    pip install xlrd==1.2.0
  ```
  3. Use python-docx to write document
  ```
    pip install python-docx
  ```


