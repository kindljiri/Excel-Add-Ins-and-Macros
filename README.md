# Excel-Add-Ins-and-Macros
How to work with Add-Ins, Create, Install, Activate see this blog post:
https://g33ksblog.blogspot.com/2019/12/excel-add-ins-with-user-defined.html

Below is list of the functions and their short description

Ip2Bin(IP)
Takes IP as input and returns binary

Bin2Ip(BinaryIP)
Takes 32 bit number and return coresponding IP (4 decimal numbers delimited by dot)

isValidIP(IP)
Returns True if the given IP is Valid IP. 4 decimal numbers in range 0-255 separated by dot

isValidMask(Mask)
Returns True if given IP is valid Mask. Valid IP which in Binary is string of consecutive 1s folowed by consecutive 0s

isValidCIDR(CIDR)
Returns True if input is valid CIDR notation IP/Mask where ex: 10.0.0.0/8

isPrivate(IP)
Returns True if IP is in private range

ReverseIP(IP)
Returns reverse IP. Ex: 140.255.15.10 -> 10.15.255.140

networkIP(IP, Mask)
returns network IP corresponding to IP and Mask combination

broadcastIP(IP, Mask)
returns broadcast IP coresponding to given IP and Mask combination

MinHost(IP, Mask)
Returns first host in given network

MaxHost(IP, Mask)
Returns last host in given network

CIDR(IP, Mask)
Returns IP and Mask in CIDR Notation. Ex: 10.0.0.0/8

CIDR2Mask(CIDR)
Returns the Mask coresponding to mask in CIDR Notation. Ex: 10.0.0.0/8 -> 255.0.0.0

Hostname(FQDN)
Return hostname of given FQDN. Ex: server1.example.com -> server1

Domain(FQDN)
Returns domain of given FQDN. Ex: server1.example.com -> example.com
