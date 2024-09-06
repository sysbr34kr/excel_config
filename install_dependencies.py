# Required imports
import subprocess
import sys
import pkg_resources

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Check if a package is installed
def check_installed(package):
    try:
        pkg_resources.require(package)
        return True
    except pkg_resources.DistributionNotFound:
        return False

# Read requirements from the file and install missing ones
def install_requirements():
    try:
        with open('requirements.txt', 'r') as f:
            packages = f.readlines()
            for package in packages:
                package = package.strip()
                if package:  # Skip empty lines
                    if check_installed(package):
                        print(f"{package} is already installed.")
                    else:
                        print(f"Installing {package}...")
                        install(package)
        print("All required packages are installed.")
    except FileNotFoundError:
        print("Error: 'requirements.txt' file not found.")

if __name__ == "__main__":
    install_requirements()