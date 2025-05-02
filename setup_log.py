import os
import datetime

def setup_log(output_log_path):
    """
    Creates and returns a log file handle, writing initial headers.
    """
    # Ensure output directory exists
    output_dir = os.path.dirname(output_log_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    log_file = open(output_log_path, 'w')
    log_file.write("Robot Joint Analysis Automation Log\n")
    log_file.write(f"Start Time: {datetime.datetime.now()}\n")
    log_file.write(f"Log file: {output_log_path}\n\n")
    return log_file