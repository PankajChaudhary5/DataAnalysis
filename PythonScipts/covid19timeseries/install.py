import sys
import subprocess

f = open('requirements.txt','r')
content = f.read()
package_list  = content.split("\n")
print(package_list)
f.close()

## implement pip as a subprocess:
for package in package_list:
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install',
        package])
    except subprocess.CalledProcessError as e:
        print(e)



## process output with an API in the subprocess module:
reqs = subprocess.check_output([sys.executable, '-m', 'pip',
'freeze'])
installed_packages = [r.decode().split('==')[0] for r in reqs.split()]

print(installed_packages)

