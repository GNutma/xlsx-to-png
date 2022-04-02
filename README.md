# .xlsx to png

For the purpose of this project, I have created a program that converts .xlsx files to .png files.
by collecting the color values of the cells in the .xlsx file, and then converting them to a .png file.

To run the program, you must have the following:

- Python >= 3.7
- an .xlsx file

the latest version of python can be found here: https://www.python.org/downloads/

put the file in the folder of the program, and then run the program within the directory. with the following command:

```shell

pip install -r requirements.txt

main.py -f <file_name> -o <output_name>
```

where: <file_name> is the name of the .xlsx file.
where:<output_name> is the name of the output file.

output_name is an optional argument. if you do not specify the output_name, the program will create a file with the same name as the input file, but with the extension .png.
