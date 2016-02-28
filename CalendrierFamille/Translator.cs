using System.Collections.Generic;
using System.Linq;

namespace CalendrierFamille
{
    public static class Translator
    {
        private static readonly Dictionary<string, string> _map = new Dictionary<string, string>()
        {
            { "Monday", "Lundi" },
            { "Tuesday", "Mardi" },
            { "Wednesday", "Mercredi" },
            { "Thursday", "Jeudi" },
            { "Friday", "Vendredi" },
            { "Saturday", "Samedi" },
            { "Sunday", "Dimanche" },
            { "Januaray", "Janvier" },
            { "February", "Février" },
            { "March", "Mars" },
            { "April", "Avril" },
            { "May", "Mai" },
            { "June", "Juin" },
            { "July", "Juillet" },
            { "August", "Août" },
            { "September", "Septembre" },
            { "Ocrober", "Octobre" },
            { "November", "Novembre" },
            { "December", "Décembre" }
        }; 

        public static string Translate(this string english)
        {
            return _map.Keys.Aggregate(english, (current, map) => current.Replace(map, _map[map]));
        }
    }
}