using System;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Management.Automation;
using Microsoft.SharePoint;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections;
using System.Collections.Generic;

namespace ExchangeAddIpAdressToConnector.VisualWebPart1
{
    public partial class VisualWebPart1UserControl : UserControl
    {
        string script;
        PowerShell ps;
        
        protected void Page_Load(object sender, EventArgs e)
        {
        }
        protected void button_ReceiveConnectors(object sender, EventArgs e)
        {
            Page.Server.ScriptTimeout = 3600; // specify the timeout to 3600 seconds
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                ListBoxConnectors.Items.Clear();
                    using (ps = PowerShell.Create())
                    {   
                        script = @"
                        $password = ConvertTo-SecureString 'Password' -AsPlainText -Force
                        $cred = New-Object System.Management.Automation.PSCredential('Username', $password)
                        $SessionOptions = New-PSSessionOption –SkipCACheck –SkipCNCheck –SkipRevocationCheck;
                        $RemoteExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://owa.ru/PowerShell/ -Authentication Basic -Credential $cred -SessionOption $SessionOptions -ErrorAction:SilentlyContinue; 
                        Import-PSSession $RemoteExSession -AllowClobber -CommandName Get-ReceiveConnector
                        $ReceiveConnectors = Get-ReceiveConnector -Server " + ListBoxMailboxServers.SelectedItem.ToString() + @" | select name,Bindings
                        #$ReceiveConnectors 
			            $arrayReceiveConnectors =@()
                        $receiveConnector = @()
			            $ReceiveConnectors | % {
                            $receiveConnector = $_.name + ', ' + $_.Bindings;
				            $Object = New-Object PSObject
				            $Object | add-member Noteproperty -Name 'name' -Value $receiveConnector -Force
				            $arrayReceiveConnectors += $Object
                        }
                        cls
                        $arrayReceiveConnectors | select -ExpandProperty name";
                        var builder = new StringBuilder();
                        PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                        ps.AddScript(script);
                        IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                        ps.EndInvoke(invocation);
                        ps.Stop();
                        foreach (PSObject outp in output)
                        {
                            if (!outp.BaseObject.ToString().Contains("tmp"))
                            {
                                ListBoxConnectors.Items.Add(outp.BaseObject.ToString());
                            }

                        }
                        ResultBox.Text += builder.ToString();
                    }

                
            }
        }
        protected void button_ReceiveIPAddress(object sender, EventArgs e)
        {
            Page.Server.ScriptTimeout = 3600; // specify the timeout to 3600 seconds
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                ListBoxIPAddress.Items.Clear();
                string stringConnector = Regex.Replace(ListBoxConnectors.SelectedItem.ToString(), @",.*", "");
                using (ps = PowerShell.Create())
                    {
                        script = @"
                        $password = ConvertTo-SecureString 'Password' -AsPlainText -Force
                        $cred = New-Object System.Management.Automation.PSCredential('Username', $password)
                        $SessionOptions = New-PSSessionOption –SkipCACheck –SkipCNCheck –SkipRevocationCheck;
                        $RemoteExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://owa.ru/PowerShell/ -Authentication Basic -Credential $cred -SessionOption $SessionOptions -ErrorAction:SilentlyContinue; 
                        Import-PSSession $RemoteExSession -AllowClobber -CommandName Get-ReceiveConnector
                        
                        $identityConnector = '" + ListBoxMailboxServers.SelectedItem.ToString() + @"' + '\' + '" + stringConnector + @"';
				        $ReceiveConnector = Get-ReceiveConnector -Identity $identityConnector;
                        cls				        
                        $ReceiveConnector.RemoteIPRanges";
                        var builder = new StringBuilder();
                        PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                        ps.AddScript(script);
                        IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                        ps.EndInvoke(invocation);
                        ps.Stop();
                        foreach (PSObject outp in output)
                        {
                            if (!outp.BaseObject.ToString().Contains("tmp"))
                            {
                                ListBoxIPAddress.Items.Add(outp.BaseObject.ToString());
                            }

                        }
                        ResultBox.Text += builder.ToString();
                    }
            }
        }
        protected void button_AddIP(object sender, EventArgs e)
        {
            Page.Server.ScriptTimeout = 3600; // specify the timeout to 3600 seconds
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                ResultBox.Text = string.Empty;
                var builder = new StringBuilder();
                string checkListBoxConnectors = string.Empty;
                for (int i = 0; i < ListBoxConnectors.Items.Count; i++)
                {
                    if (ListBoxConnectors.Items[i].Selected)
                    {
                        checkListBoxConnectors += ListBoxConnectors.Items[i].Text;
                    }
                }
                if (!string.IsNullOrEmpty(TextBox_UserName.Text))
                {
                    if (!string.IsNullOrEmpty(TextBox_Password.Text))
                    {
                        if (!string.IsNullOrEmpty(TextBox_IPRange.Text))
                        {
                            if (!string.IsNullOrEmpty(checkListBoxConnectors))
                            {
                                string stringConnector = Regex.Replace(ListBoxConnectors.SelectedItem.ToString(), @",.*", "");
                                List<string> arrayListBoxMailboxServers = new List<string>();
                                for (int i = 0; i < ListBoxMailboxServers.Items.Count; i++)
                                {
                                        arrayListBoxMailboxServers.Add(ListBoxMailboxServers.Items[i].Text);
                                }
                                string stringListBoxMailboxServers = string.Join(",", arrayListBoxMailboxServers);
                                script = @"
                                    Write-Output '++++++ START ++++++'
                                    $password = ConvertTo-SecureString '" + TextBox_Password.Text + @"' -AsPlainText -Force
                                    $cred = New-Object System.Management.Automation.PSCredential('" + TextBox_UserName.Text + @"', $password)
                                    $SessionOptions = New-PSSessionOption –SkipCACheck –SkipCNCheck –SkipRevocationCheck;
                                    $RemoteExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://owa.ru/PowerShell/ -Authentication Basic -Credential $cred -SessionOption $SessionOptions; 
                                    Import-PSSession $RemoteExSession -AllowClobber -CommandName Get-ReceiveConnector, Set-ReceiveConnector
                                    if(!$?){Write-Output 'ERROR to Connect to Exchange'; $error; break}

                                    $MailboxServers = ('" + stringListBoxMailboxServers + @"').split(',') | %{$_.trim()}
                                    $MailboxServers | %{
				                        $identityConnector = $_ + '\' + '" + stringConnector + @"';
				                        $ReceiveConnector = Get-ReceiveConnector -Identity $identityConnector;
                                        [array]$RemoteIPRanges = ('" + TextBox_IPRange.Text + @"').split(',') | %{$_.trim()}
				                        $ReceiveConnector.RemoteIPRanges += $RemoteIPRanges;
                                        cls
                                        Set-ReceiveConnector $identityConnector -RemoteIPRanges $ReceiveConnector.RemoteIPRanges
                                        if($?){
                                            Write-Output ""IP Address's range - '" + TextBox_IPRange.Text + @"' was added on the server - $_, connector - "+ stringConnector + @".""
                                        }else{
                                            Write-Output 'ERROR to set-receiveConnector on the server - $_.'
                                            $error
                                        }
                                    }
                                    #if($error){$error}
                                    Write-Output '++++++ END ++++++'
                                    #$ReceiveConnector.RemoteIPRanges
                                    ";
                                    using (ps = PowerShell.Create())
                                    {
                                        PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                        ps.AddScript(script);
                                        IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                        ps.EndInvoke(invocation);
                                        ps.Stop();
                                        foreach (PSObject outp in output)
                                        {
                                            if (!outp.BaseObject.ToString().Contains("tmp"))
                                            {
                                                ResultBox.Text += outp.ToString() + "\r\n";
                                            }
                                        }
                                    }
                                button_ReceiveIPAddress(null, null);
                                ResultBox.Text += builder.ToString();
                                
                            }
                            else { ResultBox.Text = "Please, select connector."; }
                        }
                        else { ResultBox.Text = "Please, input IP Address's range."; }
                    }
                    else { ResultBox.Text = "Please, input your Password."; }
                }
                else { ResultBox.Text = "Please, input your UserName."; }

            }
        }
        protected void button_RemoveIP(object sender, EventArgs e)
        {
            Page.Server.ScriptTimeout = 3600; // specify the timeout to 3600 seconds
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                ResultBox.Text = string.Empty;
                var builder = new StringBuilder();
                string checkListBoxConnectors = string.Empty;
                for (int i = 0; i < ListBoxConnectors.Items.Count; i++)
                {
                    if (ListBoxConnectors.Items[i].Selected)
                    {
                        checkListBoxConnectors += ListBoxConnectors.Items[i].Text;
                    }
                }
                if (!string.IsNullOrEmpty(TextBox_UserName.Text))
                {
                    if (!string.IsNullOrEmpty(TextBox_Password.Text))
                    {
                        if (!string.IsNullOrEmpty(TextBox_IPRange.Text))
                        {
                            if (!string.IsNullOrEmpty(checkListBoxConnectors))
                            {
                                string stringConnector = Regex.Replace(ListBoxConnectors.SelectedItem.ToString(), @",.*", "");
                                List<string> arrayListBoxMailboxServers = new List<string>();
                                for (int i = 0; i < ListBoxMailboxServers.Items.Count; i++)
                                {
                                        arrayListBoxMailboxServers.Add(ListBoxMailboxServers.Items[i].Text);
                                }
                                string stringListBoxMailboxServers = string.Join(",", arrayListBoxMailboxServers);
                                script = @"
                                    Write-Output '++++++ START ++++++'
                                    $password = ConvertTo-SecureString '" + TextBox_Password.Text + @"' -AsPlainText -Force
                                    $cred = New-Object System.Management.Automation.PSCredential('" + TextBox_UserName.Text + @"', $password)
                                    $SessionOptions = New-PSSessionOption –SkipCACheck –SkipCNCheck –SkipRevocationCheck;
                                    $RemoteExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://owa.ru/PowerShell/ -Authentication Basic -Credential $cred -SessionOption $SessionOptions; 
                                    Import-PSSession $RemoteExSession -AllowClobber -CommandName Get-ReceiveConnector, Set-ReceiveConnector
                                    if(!$?){Write-Output 'ERROR Connect to Exchange'}

                                    $MailboxServers = ('" + stringListBoxMailboxServers + @"').split(',') | %{$_.trim()}
                                    $MailboxServers | %{
				                        $identityConnector = $_ + '\' + '" + stringConnector + @"';
				                        $ReceiveConnector = Get-ReceiveConnector -Identity $identityConnector;
                                        [array]$RemoteIPRanges = ('" + TextBox_IPRange.Text + @"').split(',') | %{$_.trim()}
				                        $RemoteIPRanges | %{$ReceiveConnector.RemoteIPRanges.Remove($_)}
                                        Set-ReceiveConnector $identityConnector -RemoteIPRanges $ReceiveConnector.RemoteIPRanges
                                        if($?){
                                            Write-Output ""IP Address's range - '" + TextBox_IPRange.Text + @"' was removed from the server - $_, connector - " + stringConnector + @".""
                                        }else{
                                            Write-Output 'ERROR to set-receiveConnector'
                                        }
                                    }
                                    if($error){$error}
                                    Write-Output '++++++ END ++++++'
                                    #$ReceiveConnector.RemoteIPRanges
                                    ";
                                using (ps = PowerShell.Create())
                                {
                                    PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                    ps.AddScript(script);
                                    IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                    ps.EndInvoke(invocation);
                                    ps.Stop();
                                    foreach (PSObject outp in output)
                                    {
                                        if (!outp.BaseObject.ToString().Contains("tmp"))
                                        {
                                            ResultBox.Text += outp.ToString() + "\r\n";
                                        }
                                    }
                                }
                                button_ReceiveIPAddress(null, null);
                                ResultBox.Text += builder.ToString();

                            }
                            else { ResultBox.Text = "Please, select connector."; }
                        }
                        else { ResultBox.Text = "Please, input IP Address's range."; }
                    }
                    else { ResultBox.Text = "Please, input your Password."; }
                }
                else { ResultBox.Text = "Please, input your UserName."; }

            }
        }
    }
}
