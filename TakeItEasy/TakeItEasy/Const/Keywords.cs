using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TakeItEasy.Const
{
    class Keywords
    {
        public Dictionary<string, int> ListKeyword = new Dictionary<string, int>()
        {
            { "COUNT ", 0 },
            { "HAVING ", 1 },
            { "CREATE TABLE ", 2 },
            { "CREATE INDEX ", 3 },
            { "DESC ", 5 },
            { "TRUNCATE TABLE ", 6 },
            { "ALTER TABLE ", 7 },
            { "ALTER TABLE ", 8 },
            { "INSERT INTO ", 9 },
            { "UPDATE ", 10 },
            { "DELETE ", 11 },
            { "CREATE DATABASE ", 12 },
            { "DROP DATABASE ", 13 },
            { "USE ", 14 },
            { "COMMIT ", 15 },
            { "ROLLBACK ", 16 },
            { "SELECT ", 17 },
            { "DISTINCT ", 18 },
            { "WHERE ", 19 },
            { "AND ", 20 },
            { "OR ", 21 },
            { "IN ", 22 },
            { "BETWEEN ", 23 },
            { "LIKE ", 24 },
            { "ORDER BY ", 25 },
            { "GROUP BY ", 26 }
        };
    }
}
