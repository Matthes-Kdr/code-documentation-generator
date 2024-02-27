# -*- coding: utf-8 -*-
'''
Created on: Fri, 2024-02-09 (07:51:25)

@author: Matthias Kader


- Call of the main script to generate automated documentation output.
- Read the generated markdown-files:
    - "current_test.md" (the new file which has been generaditorted with the script)
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










def call_script(path:str, extra_param:str=None):
# Use subprocess to call the script
    if extra_param:
        subprocess.call(['python', path, extra_param])
    else:
        subprocess.call(['python', path])






# =============================================================================
#### MAIN: 
# =============================================================================
def main():






    # call_script('code_documenter.py')
    call_script('code_documenter.py', extra_param="--debug")

    output_path_new = "output_data/current_test.bas - Dokumentation.md"
    output_path_old = "output_data/current_test.bas - Dokumentation_REFERENCE.md"

    new_output = read_content(output_path_new)

    # old_output = read_content(output_path[:-3] + "_REFERECE.md")
    old_output = read_content(output_path_old)

    if new_output == old_output:
        print("\n\n\nOK.. No difference in {}.".format(output_path_new))
    else:
        print("\n\n\n!!!!!!!!!! DIFFERENCES TRACKT!!! please check!")



if __name__ == '__main__':
    main()