from flask import Flask, request
import subprocess
import os

app = Flask(__name__)

@app.route('/run-code', methods=['POST'])
def run_code():
    try:
        # Chạy file test.py
        result_test = subprocess.run(['python3', 'test.py'], capture_output=True, text=True)
        
        # Chạy file main.py (nếu cần chạy đồng thời)
        result_main = subprocess.run(['python3', 'main.py'], capture_output=True, text=True)

        # Kiểm tra kết quả của cả hai file
        if result_test.returncode == 0 and result_main.returncode == 0:
            return (
                f"Both main.py and test.py ran successfully.\n"
                f"Output from test.py:\n{result_test.stdout}\n"
                f"Output from main.py:\n{result_main.stdout}",
                200
            )
        else:
            return (
                f"Error occurred while processing the files.\n"
                f"Error in test.py:\n{result_test.stderr if result_test.returncode != 0 else 'No Error'}\n"
                f"Error in main.py:\n{result_main.stderr if result_main.returncode != 0 else 'No Error'}",
                500
            )

    except Exception as e:
        return f"An error occurred: {str(e)}", 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))