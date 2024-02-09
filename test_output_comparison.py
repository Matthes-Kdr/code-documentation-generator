# -*- coding: utf-8 -*-
'''
Created on: Fri, 2024-02-09 (07:51:25)

@author: Matthias Kader


- Call of the main script to generate automated documentation output.
- Read the generated markdown-files:
    - "current_test.md" (the new file which has been generated with the script)
    - "current_test_REFERENCE.md" (the file which was generated with v0.8.0+ before refactoring to syntax classes)
- compare the content of the files (except the first and last lines)
- show result (OK / DIFFERENCE)


'''












import subprocess



def read_content(path:str) -> list:

    with open(path, "r") as file:
        output_content = file.readlines()

        # dates of creation will differ:
        output_content = output_content[10:-56]

    return output_content










def call_script(path:str):
# Use subprocess to call the script
    subprocess.call(['python', path])









call_script('code_documenter.py')

output_path = "output_data/current_test.md"

new_output = read_content(output_path)

old_output = read_content(output_path[:-3] + "_REFERECE.md")

if new_output == old_output:
    print("OK.. No difference in {}.".format(new_output))
else:
    print("!!!!!!!!!! DIFFERENCES TRACKT!!! please check!")
