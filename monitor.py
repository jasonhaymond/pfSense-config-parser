import os
import subprocess
import pypandoc
import tempfile
import logging
import time
from logging.handlers import TimedRotatingFileHandler
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

LOG_DIR = os.path.dirname(os.path.abspath(__file__))  # Directory where log files are stored

# Set the working directory to the folder where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# Set up logging with handlers for both console and file
def setup_logging():
    logger = logging.getLogger("FileMonitor")
    logger.setLevel(logging.DEBUG)  # Set to DEBUG for more detailed logging

    # Create a handler that writes log messages to a file and rotates every 30 days
    log_file = os.path.join(LOG_DIR, "monitor.log")
    file_handler = TimedRotatingFileHandler(log_file, when="D", interval=30, backupCount=1)
    file_handler.setLevel(logging.DEBUG)

    # Create a console handler that outputs to the console
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)

    # Create a logging format
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # Add the handlers to the logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger

# Log cleanup function to delete log files older than 60 days
def cleanup_old_logs(directory, days=60):
    logger.info("Starting log cleanup process...")
    now = time.time()
    cutoff = now - (days * 86400)  # 86400 seconds in a day

    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        if filename.endswith(".log"):
            file_mtime = os.path.getmtime(file_path)
            if file_mtime < cutoff:
                try:
                    os.remove(file_path)
                    logger.info(f"Deleted old log file: {file_path}")
                except Exception as e:
                    logger.error(f"Failed to delete log file {file_path}: {e}")

# Initialize the logger
logger = setup_logging()

logger.info("----- Starting Monitor -----")

class FileCreationHandler(FileSystemEventHandler):
    def __init__(self):
        pass

    def on_created(self, event):
        logger.debug(f"New file detected: {event.src_path}")
        if not event.is_directory:
            # Get the full path and filename of the new file
            file_path = event.src_path
            filename = os.path.basename(file_path)
            directory_path = os.path.dirname(file_path)

            # Exclude .docx files
            if filename.lower().endswith('.docx'):
                return
            
            logger.debug(f"File created: {filename} in directory: {directory_path}")
            
            # Check if the new file is an XML file and starts with 'config-'
            if filename.lower().endswith('.xml') and filename.startswith('config-'):
                logger.info(f"Processing new XML file: {filename}")
                try:
                    # Use a temporary file for the .md file
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.md') as temp_md_file:
                        md_file_path = temp_md_file.name
                        md_filename = os.path.basename(md_file_path)
                        logger.info(f"Temporary .md file created: {md_file_path}")
                        
                        # Process the new file
                        self.process_new_file(directory_path, file_path, md_filename, md_file_path)
                except Exception as e:
                    logger.error(f"Error while creating temporary .md file: {e}")

    def process_new_file(self, directory_path, file_path, md_filename, md_file_path):
        logger.info(f"Starting processing of file: {file_path}")
        
        # Run the external command to create the .md file
        command = f"pf-format -i '{file_path}' -f md -o '{md_file_path}'"
        logger.debug(f"Running command: {command}")
        try:
            result = subprocess.run(['powershell', '-command', command], capture_output=True, text=True, encoding='utf-8')
            logger.info(f"Command output: {result.stdout}")
            if result.stderr:
                logger.error(f"Command errors: {result.stderr}")
            logger.info(f"Processing '{md_filename}' complete.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Error while running pf-format: {e}")
            return  # Exit the function if the command fails

        # Convert the .md file to a .docx file
        logger.info(f"Converting MD file to DOCX for: {md_filename}")
        docx_filename = os.path.splitext(file_path)[0] + '.docx'
        docx_file_path = os.path.join(directory_path, docx_filename)

        # Ensure the output directory exists
        try:
            docx_directory = os.path.dirname(docx_file_path)
            if not os.path.exists(docx_directory):
                os.makedirs(docx_directory)
                logger.info(f"Created missing directory: {docx_directory}")
        except OSError as e:
            logger.error(f"Error creating directory {docx_directory}: {e}")
            return  # Exit the function if the directory creation fails

        # Convert the temporary Markdown file to DOCX
        try:
            output = pypandoc.convert_file(md_file_path, 'docx', outputfile=docx_file_path)
            if output == "":
                logger.info(f"Conversion to DOCX successful: {docx_file_path}")
            else:
                logger.error(f"Error in conversion: {output}")
        except Exception as e:
            logger.error(f"Error during conversion: {e}")
            return  # Exit the function if the conversion fails

        # Delete the temporary .md file after successful conversion
        try:
            os.remove(md_file_path)
            logger.info(f"Temporary .md file deleted: {md_file_path}")
        except OSError as e:
            logger.error(f"Error deleting temporary .md file: {e}")

        logger.info(f"Monitoring folder...")

def monitor_directory(paths_to_watch):
    """Monitors one or more directories for new files."""
    event_handler = FileCreationHandler()
    observer = Observer()

    # Schedule each directory for monitoring
    for path in paths_to_watch:
        try:
            observer.schedule(event_handler, path=path, recursive=True)
            logger.info(f"Monitoring directory: {path} (recursive)")
        except OSError as e:
            logger.error(f"Error monitoring directory {path}: {e}")

    observer.start()

    try:
        while True:
            pass  # Keep the script running
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    # Set the directories to monitor (replace with paths as needed)
    paths_to_watch = [
        # Add directories here separated by commas (one per line).
    ]

    if not paths_to_watch:
        # Prompt user for paths one by one, press Enter with no input to stop
        logger.info("Enter directories to monitor (one per line). Press Enter on a blank line to finish:")
        while True:
            input_path = input("Enter path: ").strip()
            if input_path:
                paths_to_watch.append(input_path)
            else:
                break

    # If no directories were entered, exit with an error message
    if not paths_to_watch:
        logger.error("No directories provided to monitor. Exiting.")
        exit(1)  # Exit with a non-zero status to indicate an error

    # Start monitoring the directories recursively
    monitor_directory(paths_to_watch)

    # Run log cleanup (deletes logs older than 60 days)
    cleanup_old_logs(LOG_DIR, days=60)
