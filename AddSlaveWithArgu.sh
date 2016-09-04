#!/bin/bash
domain=("$@")
echo "zone \"$domain\" { type slave; file \"/var/named/slaves/$domain\"; masters { <1st master IP>; <2nd master IP>; }; };" >> /var/named/chroot/etc/named.zone
exit 0;
