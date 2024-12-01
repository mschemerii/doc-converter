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

def activate_venv():
    """Activate the virtual environment and return the activated environment path"""
    venv_path = Path('.venv')
    
    if sys.platform == 'win32':
        python_path = venv_path / 'Scripts' / 'python'
        pip_path = venv_path / 'Scripts' / 'pip'
        site_packages = venv_path / 'Lib' / 'site-packages'
    else:
        python_path = venv_path / 'bin' / 'python'
        pip_path = venv_path / 'bin' / 'pip'
        site_packages = venv_path / 'lib' / f'python{sys.version_info.major}.{sys.version_info.minor}' / 'site-packages'
    
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
    
    return env, str(python_path), str(pip_path)

def install_requirements(pip_path, env):
    """Install packages from requirements.txt"""
    try:
        print("Installing required packages...")
        result = subprocess.run(
            [pip_path, 'install', '-r', 'requirements.txt'],
            env=env,
            check=True,
            capture_output=True,
            text=True
        )
        print(result.stdout)
        if result.stderr:
            print("Warnings:", result.stderr)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error installing requirements: {e.stderr}")
        return False

def parse_requirements(filename='requirements.txt'):
    """Parse requirements file and filter based on platform compatibility"""
    if not os.path.exists(filename):
        print(f"Warning: {filename} not found. Installing only python-docx.")
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

def create_venv():
    """Create and activate a Python virtual environment"""
    venv_path = Path('.venv')
    
    # Create the virtual environment
    print("Creating virtual environment...")
    venv.create(venv_path, with_pip=True)
    
    # Activate the environment and get the new environment variables
    env, python_path, pip_path = activate_venv()
    
    # Install requirements
    if not install_requirements(pip_path, env):
        return False
    
    print("Virtual environment created, activated, and packages installed.")
    return True

def check_venv():
    """Check if running in a virtual environment and create one if not present"""
    venv_path = Path('.venv')
    
    if venv_path.exists():
        print("\nExisting virtual environment found.")
        # Activate the environment and get the new environment variables
        env, python_path, pip_path = activate_venv()
        
        # Ensure requirements are installed
        if not install_requirements(pip_path, env):
            return False
        
        print("Virtual environment activated and packages verified.")
        return True
    else:
        print("\nNo virtual environment detected.")
        print("Creating virtual environment automatically...")
        if not create_venv():
            return False
        return True

def run_script(script_name, input_file, description):
    """Run a Python script and handle its output"""
    print(f"\n=== {description} ===")
    try:
        # Get the Python interpreter from the virtual environment if it exists
        python_cmd = str(Path('.venv') / ('Scripts' if sys.platform == 'win32' else 'bin') / 'python')
        if not os.path.exists(python_cmd):
            python_cmd = 'python3'
        
        result = subprocess.run(
            [python_cmd, script_name, input_file],
            capture_output=True,
            text=True,
            check=True,
            env=os.environ
        )
        print(result.stdout)
        if result.stderr:
            print("Warnings:", result.stderr)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error running {script_name}:")
        print(e.stderr)
        return False

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
        
        print("\nVirtual environment deactivated.")

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
        print("Missing required Linux dependencies:", ', '.join(missing_deps))
        print("\nPlease install them using your package manager:")
        print("\nFor Ubuntu/Debian:")
        print("sudo apt-get update")
        print(f"sudo apt-get install {' '.join(missing_deps)}")
        print("\nFor Oracle Linux/RHEL:")
        print("sudo yum update")
        print(f"sudo yum install {' '.join(missing_deps)}")
        return False
    
    return True

def process_document(doc_path):
    """Process a .doc file through all conversion and modification steps"""
    # Check platform-specific dependencies
    if not check_linux_dependencies():
        sys.exit(1)
    
    if not os.path.exists(doc_path):
        print(f"Error: File not found: {doc_path}")
        return False
    
    # Get the directory and base filename
    directory = os.path.dirname(doc_path) or '.'
    filename = os.path.basename(doc_path)
    base_name, ext = os.path.splitext(filename)
    
    if ext.lower() != '.doc':
        print("Error: Input file must be a .doc file")
        return False
    
    # Define the intermediate docx filename
    docx_path = os.path.join(directory, f"{base_name}.docx")
    
    # Step 1: Convert .doc to .docx
    if not run_script(
        'doc_to_docx_converter.py',
        doc_path,
        "Converting .doc to .docx"
    ):
        return False
    
    # Small delay to ensure file is ready
    time.sleep(1)
    
    # Step 2: Modify table properties
    if not run_script(
        'modify_docx_tables.py',
        docx_path,
        "Modifying table properties"
    ):
        return False
    
    # Step 3: Add table rows
    if not run_script(
        'add_table_rows.py',
        docx_path,
        "Adding empty rows to tables"
    ):
        return False
    
    # Step 4: Create renamed copies with headers
    if not run_script(
        'rename_docx.py',
        docx_path,
        "Creating final copies with headers"
    ):
        return False
    
    print("\n=== Processing Complete ===")
    print(f"Original .doc file: {doc_path}")
    print(f"Intermediate .docx file: {docx_path}")
    print("Final files created with appropriate headers and content modifications.")
    return True

def main():
    if len(sys.argv) != 2:
        print("Usage: python process_document.py <path_to_doc_file>")
        sys.exit(1)
    
    try:
        # Check virtual environment first
        if not check_venv():
            sys.exit(1)
        
        doc_path = sys.argv[1]
        success = process_document(doc_path)
        
        # Deactivate virtual environment before exiting
        deactivate_venv()
        
        if success:
            print("\nDocument processing completed successfully!")
        else:
            print("\nDocument processing failed.")
            sys.exit(1)
    except Exception as e:
        # Ensure we deactivate even if there's an error
        deactivate_venv()
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()
