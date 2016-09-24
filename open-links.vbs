Const navOpenInBackgroundTab = &H1000

VKB = "https://vkbexternal.partners.extranet.microsoft.com/vkbweb/?portalid=1"
KIT = "https://kit.one.microsoft.com/"
PACSR = "https://mspacsr.microsoft.com/"
LMI = "https://aka.ms/lmi"
OfficeD3 = "https://sharepoint.partners.extranet.microsoft.com/sites/GCS/commboard/Pages/Office-Decision-Tree-PIlot.aspx"
WUE = "https://wue2.hotmail.com/default.aspx"
DR = "https://gc.digitalriver.com/gc/ent/login.do"
QC = "http://quickconnect.convergys.com/launch/"
Playbook = "https://cvgsharepoint/sites/msoffice-mdc/ADO/Shared%20Documents/Forms/AllItems.aspx"

Set oIE = CreateObject("InternetExplorer.Application")
oIE.Visible = True
oIE.Navigate2 VKB
oIE.Navigate2 KIT,navOpenInBackgroundTab
oIE.Navigate2 PACSR,navOpenInBackgroundTab
oIE.Navigate2 LMI,navOpenInBackgroundTab
oIE.Navigate2 OfficeD3,navOpenInBackgroundTab
oIE.Navigate2 WUE,navOpenInBackgroundTab
oIE.Navigate2 DR,navOpenInBackgroundTab
oIE.Navigate2 QC,navOpenInBackgroundTab
oIE.Navigate2 Playbook,navOpenInBackgroundTab

Set oIE = Nothing