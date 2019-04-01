#include <iostream>
#include <fstream>
#include <string>
#include <map>
#include <Windows.h>
#include <sstream>
#include <vector>
#include <iterator>

#include <codecvt>

using namespace std;
typedef map<string, int> tMap;
typedef multimap<int,string> tMmap;

int main()
{
	ifstream ifs;
    std::noskipws(ifs);
    ifs.open("football1.txt", ios::in | ios::binary);

	string ss = "";
	tMap m_pWordMap;
	tMap::iterator itr;

    while (!ifs.eof())
    {
        char c;
        ifs.get(c);
		
		ss += c;
		if (c == '\n')
		{
			string tmp = ss.substr(0, ss.find("\t"));
			char* pch = strtok(strdup(tmp.c_str()), " ");
			for (pch; pch != NULL; pch = strtok(NULL, " "))
			{
				if ((itr = m_pWordMap.find(pch)) != m_pWordMap.end())
				{
					itr->second++;
				}
				else
				{
					m_pWordMap.insert(pair<string, int>(pch, 1));
				}
			}
			ss = "";
		}
    }

	ifs.close();

	tMmap test2;
	for (tMap::iterator it = m_pWordMap.begin(); it != m_pWordMap.end(); ++it)
		test2.insert(pair<int, string>(it->second,it->first));

	ofstream of ("of.txt");

	for (tMmap::reverse_iterator it = test2.rbegin(); it != test2.rend(); ++it)
	{
		of << it->first << " " << it->second << endl;

		if (it->first == 100)
				break;
	}

	of.close();
}