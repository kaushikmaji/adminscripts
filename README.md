# Scripts for IT admins

This project is a collection of scripts in powershell (as of today and would be extended with SH, Python and other languages) for system administration and reporting. These are typically used by operations/DevOPs teams to ensure desired health of the scope of services and servers being managed.

## Getting Started

You can review and download scripts as necessary, follow pre-req section for respective script and esure these are met, before running scripts.

### Prerequisites

Pre-requisites are specific to each script and are mentioned atop the executable lines in each script file.

```
e.g. Powershell_Powershell_CheckClusterResourceHealth.ps1 has below pre-reqs 
	WMI should be enabled on all target servers
	WMI ports to all target servers, should be opened from the location where the script is run  
	Excel should be installed in the location where the script is run
```

### Installing

Please run the required script as per guidelines provided atop the executable lines in each script file.

 - Powershell_CheckClusterResourceHealth.ps1
	> The script generates an excel report showing health status of resources in a windows cluster
 - Powershell_DiskUtilizationReport.ps1
	> Script to obtain the disk space utilization on remote servers & failover clusters
 - Powershell_GetAllSystemInfo.ps1
	> This helps extract OS, disk, disk partition, free space on disk, CPU/RAM, Fileshares, BIOS and other information for a list of servers.


## Contributing

Please read [CONTRIBUTING.md](https://gist.github.com/PurpleBooth/b24679402957c63ec426) for details on our code of conduct, and the process for submitting pull requests to us.

## Versioning

For the versions available, see the [tags on this repository](https://github.com/kaushikmaji/adminscripts/tags). 

## Authors

* **Kaushik Maji** - *Initial work* - [kaushikmaji](https://github.com/kaushikmaji)

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Hat tip to anyone whose code was used
* Inspiration
* etc
