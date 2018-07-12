# ExTRATrace
This script starts and stops logman trace collection on one or many Exchange servers simultaneously. After collection is stopped logs are collected to local server for review.Engineers can generate ExTRA configuration strings to provide to end users to collect specific datafrom Exchange.

Exchange 2010SP3, 2013, and 2016 supported as long as compatible tags are provided.

# Usage Examples

  - Interactive Configuration generator
  
    *.\ExTRAtrace.ps1 -Generate*

  - Start ExTRA log generation after prompting for configuration
  
    *.\ExTRAtrace.ps1 -Start*

  - Stop ExTRA tracing and consolidate logs into D:\logs\extra\
  
    *.\ExTRAtrace.ps1 -Stop -LogPath "D:\logs\extra\"*

