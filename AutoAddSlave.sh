#!/bin/sh

rm -rf <delete old list>
rm -rf /var/named/chroot/etc/named.zone
wget <get list from first server>
wget <get list from second server>

array1=(`cat 1st.txt`)
array2=(`cat 2nd.txt`)

declare -A array4

for i in "${array1[@]}" "${array2[@]}"; do array4["$i"]=1; done
echo "${!array4[@]}" > domain.txt

domain=(`cat domain.txt`)
for ((i=0; i<${#domain[@]}; i++))
do
echo "zone \"${domain[i]}\" { type slave; file \"/var/named/slaves/${domain[i]}\"; masters { <1st master IP>; <2nd master IP>; }; };" >> /var/named/chroot/etc/named.zone;
done
/etc/init.d/named restart
exit 0;
