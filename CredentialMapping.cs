using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsys.EAI.Framework.PasswordManager
{
    public class Map : IEquatable<Map>
    {
        public string UriStartWith { get; set; }

        public string Username { get; set; }

        public string Password { get; set; }

        public bool Equals(Map obj)
        {
            return (UriStartWith == obj.UriStartWith && Username == obj.Username);
        }

    }

    public class CredentialMapping
    {
        public List<Map> Maps { get; set; }
    }

}
