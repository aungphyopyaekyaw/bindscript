usrlist=(`cat usrlist.txt`)
for ((i=0; i<${#usrlist[@]}; i++))
do
tpass=`openssl rand -base64 12`
echo -e "$tpass\n$tpass" | passwd ${usrlist[i]}
echo -e "${usrlist[i]} $tpass" >> passwds.txt
done
