using System;
using System.Net.Mail;
using System.DirectoryServices.AccountManagement;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Gentex.ComponentTracker.Plugin;

namespace OutlookParser
{
  public class AdEmailResolver : IEmailResolver
  {
    public string DefaultDomain { get; set; }
    private Regex _regexDomain = new Regex(@"\w+\\([-.A-Za-z0-9_]+)@[?]");

    public AdEmailResolver(string defaultDomain)
    {
      if (defaultDomain == null) throw new ArgumentException("defaultDomain");
      this.DefaultDomain = defaultDomain;
    }

    public MailAddress ProcessAddress(string displayName, string email)
    {
      if (string.IsNullOrEmpty(email))
      {
        return null;
      }
      else
      {
        var recipIndex = email.IndexOf("/CN=RECIPIENTS/CN=", StringComparison.InvariantCultureIgnoreCase);
        if (recipIndex < 0)
        {
          var match = _regexDomain.Match(email);
          if (match.Success)
          {
            using (var data = (GetDataByLegacyDn(this.DefaultDomain, "*/CN=RECIPIENTS/CN=" + match.Groups[1].Value) ?? 
                               GetDataByUserName(this.DefaultDomain, match.Groups[1].Value)))
            {
              if (data == null)
              {
                return new MailAddress("?@?", displayName ?? "");
              }
              else
              {
                return new MailAddress(data.EmailAddress ?? "?@?", data.Name ?? "");
              }
            }
          }
          else
          {
            try
            {
              return new MailAddress(email ?? "?@?", displayName ?? "");
            }
            catch (FormatException)
            {
              return null;
            }
          }
        }
        else
        {
          string emailConv;
          using (var data = GetDataByLegacyDn(this.DefaultDomain, email))
          {
            if (data == null)
            {
              var i = recipIndex + 18;
              while (i < email.Length && (char.IsLetterOrDigit(email[i]) || char.IsPunctuation(email[i]))) i++;
              emailConv = email.Substring(recipIndex + 18, i - (recipIndex + 18)).ToLowerInvariant() + "@" + this.DefaultDomain;
            }
            else
            {
              emailConv = data.EmailAddress;
            }
          }
          return new MailAddress(emailConv, displayName ?? "");
        }
      }
    }

    private static Dictionary<string, PrincipalContext> _contexts = new Dictionary<string, PrincipalContext>();
    private static PrincipalContext GetContext(string domain)
    {
      PrincipalContext result = null;
      if (!_contexts.TryGetValue(domain, out result))
      {
        result = new PrincipalContext(ContextType.Domain, domain);
        _contexts[domain] = result;
      }
      return result;
    }


    public static UserPrincipal GetDataByUserName(string domain, string name)
    {
      var pc = GetContext(domain);
      using (var user = new UserPrincipal(pc))
      {
        user.SamAccountName = name;
        using (var searcher = new PrincipalSearcher(user))
        {
          return searcher.FindAll().Select(r => (UserPrincipal)r).SingleOrDefault();
        }
      }
    }
    
    public static UserPrincipal GetDataByLegacyDn(string domain, string name)
    {
      var pc = GetContext(domain);
      using (var user = new AdvancedUserPrincipal(pc))
      {
        user.LegacyExchangeDn = name;
        using (var searcher = new PrincipalSearcher(user))
        {
          return searcher.FindAll().Select(r => (UserPrincipal)r).SingleOrDefault();
        }
      }
    }

    [DirectoryRdnPrefix("CN")]
    [DirectoryObjectClass("User")]
    public class AdvancedUserPrincipal : UserPrincipal
    {

      public AdvancedUserPrincipal(PrincipalContext context)
        : base(context)
      {
      }
      public AdvancedUserPrincipal(PrincipalContext context, string samAccountName, string password, bool enabled)
        : base(context, samAccountName, password, enabled)
      {
      }

      [DirectoryProperty("legacyExchangeDN")]
      public string LegacyExchangeDn
      {
        get
        {
          var value = ExtensionGet("legacyExchangeDN");
          if (value.Length != 1)
            return null;
          return value[0].ToString();
        }
        set { ExtensionSet("legacyExchangeDN", value); }
      }
    }
  }
}
