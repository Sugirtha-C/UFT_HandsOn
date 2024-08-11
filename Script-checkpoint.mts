Browser("OrangeHRM").Page("OrangeHRM").WebEdit("username").Set "admin" @@ script infofile_;_ZIP::ssf1.xml_;_
Browser("OrangeHRM").Page("OrangeHRM").WebEdit("password").SetSecure "66b18bae3713eedf1bcbc3a78fd197c9a990c24b60ab" @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("OrangeHRM").Page("OrangeHRM").WebButton("Login").Check CheckPoint("Login") @@ script infofile_;_ZIP::ssf3.xml_;_
