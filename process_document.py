#!/usr/bin/env python3
import sys
import os
import subprocess
import time
import venv
from pathlib import Path
import platform
import site
import importlib
import logging
import traceback

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler('doc_converter.log', mode='a'),
        logging.StreamHandler(sys.stdout)
    ]
)

def activate_venv():
    """Activate the virtual environment and return the activated environment path"""
    try:
        venv_path = Path('.venv')
        
        if sys.platform == 'win32':
            python_path = venv_path / 'Scripts' / 'python'
            pip_path = venv_path / 'Scripts' / 'pip'
            site_packages = venv_path / 'Lib' / 'site-packages'
        else:
            python_path = venv_path / 'bin' / 'python'
            pip_path = venv_path / 'bin' / 'pip'
            site_packages = venv_path / 'lib' / f'python{sys.version_info.major}.{sys.version_info.minor}' / 'site-packages'
        
        # Validate virtual environment exists
        if not venv_path.exists():
            logging.error(f"Virtual environment not found at {venv_path}")
            raise FileNotFoundError(f"Virtual environment not found at {venv_path}")
        
        # Add virtual environment to Python path
        if str(site_packages) not in sys.path:
            sys.path.insert(0, str(site_packages))
        
        # Create new environment with venv path
        env = os.environ.copy()
        env['VIRTUAL_ENV'] = str(venv_path)
        env['PATH'] = str(venv_path / ('Scripts' if sys.platform == 'win32' else 'bin')) + os.pathsep + env['PATH']
        
        # Unset PYTHONHOME if set
        env.pop('PYTHONHOME', None)
        
        # Reload site module to recognize the new environment
        importlib.reload(site)
        
        logging.info(f"Virtual environment activated at {venv_path}")
        return env, str(python_path), str(pip_path)
    
    except Exception as e:
        logging.error(f"Error activating virtual environment: {e}")
        logging.error(traceback.format_exc())
        raise

def install_requirements(pip_path, env):
    """Install packages from requirements.txt"""
    try:
        logging.info("Installing required packages...")
        result = subprocess.run(
            [pip_path, 'install', '-r', 'requirements.txt'],
            env=env,
            check=True,
            capture_output=True,
            text=True
        )
        logging.info("Successfully installed requirements")
        return result
    except subprocess.CalledProcessError as e:
        logging.error(f"Error installing requirements: {e.stderr}")
        logging.error(traceback.format_exc())
        raise

def create_venv(venv_path=Path('.venv')):
    """Create a new virtual environment"""
    try:
        if venv_path.exists():
            logging.warning(f"Virtual environment already exists at {venv_path}")
            return
        
        logging.info(f"Creating virtual environment at {venv_path}")
        venv.create(venv_path, with_pip=True)
        
        logging.info("Virtual environment created successfully")
    except Exception as e:
        logging.error(f"Error creating virtual environment: {e}")
        logging.error(traceback.format_exc())
        raise

def parse_requirements(filename='requirements.txt'):
    """Parse requirements file and filter based on platform compatibility"""
    if not os.path.exists(filename):
        logging.warning(f"{filename} not found. Installing only python-docx.")
        return ['python-docx']
    
    requirements = []
    current_platform = platform.system().lower()
    
    with open(filename, 'r') as f:
        for line in f:
            # Skip comments and empty lines
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            
            # Check for platform-specific markers
            if ';' in line:
                pkg, marker = line.split(';', 1)
                pkg = pkg.strip()
                marker = marker.strip()
                
                # Basic platform checking
                if 'platform_system' in marker:
                    if current_platform == 'darwin' and 'darwin' not in marker.lower():
                        continue
                    if current_platform == 'windows' and 'windows' not in marker.lower():
                        continue
                    if current_platform == 'linux' and 'linux' not in marker.lower():
                        continue
                    requirements.append(pkg)
                else:
                    requirements.append(pkg)
            else:
                requirements.append(line)
    
    return requirements

def check_linux_dependencies():
    """Check if required Linux dependencies are installed"""
    system = platform.system()
    if system != 'Linux':
        return True

    missing_deps = []
    
    # Check for LibreOffice
    if not any(os.path.exists(path) for path in [
        '/usr/bin/soffice',
        '/usr/lib/libreoffice/program/soffice'
    ]):
        missing_deps.append('libreoffice')
    
    # Check for unoconv
    try:
        subprocess.run(['unoconv', '--version'], capture_output=True, check=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        missing_deps.append('unoconv')
    
    if missing_deps:
        logging.error("Missing required Linux dependencies: " + ', '.join(missing_deps))
        logging.error("\nPlease install them using your package manager:")
        logging.error("\nFor Ubuntu/Debian:")
        logging.error("sudo apt-get update")
        logging.error(f"sudo apt-get install {' '.join(missing_deps)}")
        logging.error("\nFor Oracle Linux/RHEL:")
        logging.error("sudo yum update")
        logging.error(f"sudo yum install {' '.join(missing_deps)}")
        return False
    
    return True

def process_document(doc_path):
    """Process a .doc file through all conversion and modification steps"""
    try:
        logging.info(f"Starting document processing for {doc_path}")
        
        # Validate input file
        if not os.path.exists(doc_path):
            logging.error(f"Input file not found: {doc_path}")
            raise FileNotFoundError(f"Input file not found: {doc_path}")
        
        # Get absolute paths
        doc_path = os.path.abspath(doc_path)
        docx_path = os.path.splitext(doc_path)[0] + '.docx'
        
        # Get the current script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Get the Python executable path
        if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
            python_cmd = sys.executable
        else:
            python_cmd = 'python3'
        
        logging.info(f"Using Python executable: {python_cmd}")
        
        # Change to script directory
        original_dir = os.getcwd()
        os.chdir(script_dir)
        
        try:
            # Step 1: Convert .doc to .docx
            logging.info("Starting .doc to .docx conversion...")
            result = subprocess.run(
                [python_cmd, 'doc_to_docx_converter.py', doc_path],
                capture_output=True,
                text=True,
                check=True,
                env=os.environ
            )
            logging.info(result.stdout)
            if result.stderr:
                logging.warning("Conversion warnings: " + result.stderr)
            
            # Small delay to ensure file is ready
            time.sleep(1)
            
            # Step 2: Modify table properties
            result = subprocess.run(
                [python_cmd, 'modify_docx_tables.py', docx_path],
                capture_output=True,
                text=True,
                check=True,
                env=os.environ
            )
            logging.info(result.stdout)
            if result.stderr:
                logging.warning("Warnings: " + result.stderr)
            
            # Step 3: Add table rows
            result = subprocess.run(
                [python_cmd, 'add_table_rows.py', docx_path],
                capture_output=True,
                text=True,
                check=True,
                env=os.environ
            )
            logging.info(result.stdout)
            if result.stderr:
                logging.warning("Warnings: " + result.stderr)
            
            # Step 4: Create renamed copies with headers
            result = subprocess.run(
                [python_cmd, 'rename_docx.py', docx_path],
                capture_output=True,
                text=True,
                check=True,
                env=os.environ
            )
            logging.info(result.stdout)
            if result.stderr:
                logging.warning("Warnings: " + result.stderr)
            
            logging.info("\n=== Processing Complete ===")
            logging.info(f"Original .doc file: {doc_path}")
            logging.info(f"Intermediate .docx file: {docx_path}")
            logging.info("Final files created with appropriate headers and content modifications.")
            return True
        
        finally:
            # Restore original directory
            os.chdir(original_dir)
            
        return True
    
    except subprocess.CalledProcessError as e:
        logging.critical(f"Command failed with exit status {e.returncode}")
        logging.critical(f"Command output: {e.stdout}")
        logging.critical(f"Command error: {e.stderr}")
        raise
    except Exception as e:
        logging.critical(f"Document processing failed: {e}")
        logging.critical(traceback.format_exc())
        raise

def deactivate_venv():
    """Deactivate the virtual environment"""
    if 'VIRTUAL_ENV' in os.environ:
        # Remove virtual environment from PATH
        path = os.environ['PATH'].split(os.pathsep)
        venv_path = os.environ['VIRTUAL_ENV']
        path = [p for p in path if not p.startswith(venv_path)]
        os.environ['PATH'] = os.pathsep.join(path)
        
        # Remove virtual environment from sys.path
        if str(Path(venv_path) / 'lib' / f'python{sys.version_info.major}.{sys.version_info.minor}' / 'site-packages') in sys.path:
            sys.path.remove(str(Path(venv_path) / 'lib' / f'python{sys.version_info.major}.{sys.version_info.minor}' / 'site-packages'))
        
        # Unset virtual environment variables
        del os.environ['VIRTUAL_ENV']
        if 'PYTHONHOME' in os.environ:
            del os.environ['PYTHONHOME']
        
        # Reload site module to recognize the changes
        importlib.reload(site)
        
        logging.info("\nVirtual environment deactivated.")

def main():
    """Main entry point for document processing"""
    try:
        # Check for input file
        if len(sys.argv) < 2:
            logging.error("Usage: python process_document.py <input_file.doc>")
            sys.exit(1)
        
        input_file = sys.argv[1]
        
        # Create and activate virtual environment
        create_venv()
        env, python_path, pip_path = activate_venv()
        
        # Install requirements
        install_requirements(pip_path, env)
        
        # Process document
        output_file = process_document(input_file)
        print(f"Successfully processed document: {output_file}")
        
        # Deactivate virtual environment
        deactivate_venv()
    
    except Exception as e:
        logging.error(f"Processing failed: {e}")
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
