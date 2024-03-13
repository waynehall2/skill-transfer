# skill-transfer
Transfer employee skills to different Excel format

download excel doc and decrypt it
assumes bash/linux (ubuntu specifically)
install pandas and openpyxl on your local machine
```bash
sudo apt update # if pip not installed yet
sudo apt install python3-pip -y # if pip not installed yet
pip install pandas openpyxl
```

run the skills.py against your source file
```bash
python3 skills.py /home/user/SkillsSpreadsheet.xlsx
```
