# Peyton-Lab-Microscopy-Analysis
A tool that will take intensity data, format it by condition, timepoint, and fluorophore, and generate correlative graphs.

How to use this tool if you do not have Python 3.8 installed:

1) Install Anaconda. This is a Python distribution that will automatically install Python, a development environment, and all necessary libraries for this program.
  - Follow the installation intructions at this link: https://docs.anaconda.com/anaconda/install/
2) Make a new folder where you want the program, the program outputs, and the program inputs to reside. For instructions below, we will refer to this folder as "myfolder"
3) Download transpose_data_user_input.py and put it in this folder.
4) Put all of your raw data .csv files in the same folder. 

Your files are now prepped and ready to execute.
First, obtain the full path of myfolder and copy it. 
  - For example, it might be: mypc\Downloads\myfolder 
  - We will refer to the full path as "mypath"

To execute the program, open command prompt and type the following commands in the following order:

1) `cd mypath`
2) `python transpose_data_user_input.py`

You're all set! The program is now running on your command prompt.
