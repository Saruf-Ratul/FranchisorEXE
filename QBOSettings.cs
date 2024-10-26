using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
public class QBOSettings
{
    public string CompanyID = "";
    public string FranchisorID = "";
    public string LocationID = "";
    public string LocationName = "";

    public string UserID = "";
    public string AccessToken = "";
    public string RefreshToken = "";
    public string FileID = "";
    public string ErrorMessage = "";

    public bool IsUserValid = false;
    public bool IsQBOConnected = false;
    public bool Sandbox = false;
    public DateTime LastSyncDateTime = DateTime.Now;

    public string FirstName = "";
    public string LastName = "";
    public string ConnectedLocationName = "";
    public string QuickbookCompanyName = "";

}
