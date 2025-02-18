import pkg_resources
import sys
import platform

def print_package_info():
    print(f"Python Version: {sys.version}")
    print(f"Platform: {platform.platform()}")
    print("\nInstalled Packages:")
    print("-" * 60)
    print(f"{'Package':<30} {'Version':<15} {'Location'}")
    print("-" * 60)
    
    for pkg in sorted(pkg_resources.working_set, key=lambda x: x.key):
        print(f"{pkg.key:<30} {pkg.version:<15} {pkg.location}")

if __name__ == "__main__":
    print_package_info() 