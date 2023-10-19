import subprocess


def run_matomo_process(input_file):
    print(f"Running matomo_process.py with input {input_file}...")
    command = f"python matomo_processing.py {input_file}"
    completed_process = subprocess.run(command, shell=True)
    if completed_process.returncode == 0:
        print("matomo_process.py completed successfully.")
    else:
        print(f"matomo_process.py failed with return code {completed_process.returncode}.")


def run_qualtrics_matching(input_csv, output_xlsx):
    print(f"Running qualtrics_matching.py with inputs {input_csv} and {output_xlsx}...")
    command = f"python qualtrics_matching.py {input_csv} {output_xlsx}"
    completed_process = subprocess.run(command, shell=True)
    if completed_process.returncode == 0:
        print("qualtrics_matching.py completed successfully.")
    else:
        print(f"qualtrics_matching.py failed with return code {completed_process.returncode}.")


if __name__ == "__main__":
    matomo_input_file = 'Matomo3.csv'
    qualtrics_input_csv = 'Qualtrics6.csv'
    qualtrics_output_xlsx = 'Qualtrics_Validation_All_10-03To10-16.xlsx'

    run_matomo_process(matomo_input_file)
    run_qualtrics_matching(qualtrics_input_csv, qualtrics_output_xlsx)
