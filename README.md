# ![halo](https://cdn2.iconfinder.com/data/icons/windows-8-metro-style/48/plasmid.png)harken
Harken is a Windows Service and .NET HTTP Listener that allows you to run PowerShell commands via HTTP requests.

### Install Harken on Windows
1. **Run PowerShell as an Administrator**
2. `git clone` this repository.
3. `cd harken`
4. Source the file with `. \harken.ps1` 
5. `Install-Harken`
6. Visit: `http://localhost:8888/harken?command=get-service winmgmt&format=text`
7. JSON PLZ?: `http://localhost:8888/harken?command=get-service winmgmt&format=json`

### Uninstall Harken
1. **Run PowerShell as an Administrator**
2. `cd harken`
3. Source the file with `. \harken.ps1` 
4. `Uninstall-Harken`

#### Made Possible By:
* [PowerShell Team's Simple HTTP API Script](https://gallery.technet.microsoft.com/scriptcenter/Simple-REST-api-for-b04489f1)
* [NSSM - The Non-Sucking  Service Manager](http://www.nssm.cc/) 
* [Dill Pickle Spitz Sunflower Seeds](http://www.fritolay.com/snacks/product-page/nuts-and-seeds/spitz-dill-pickle-flavored-sunflower-seeds)
