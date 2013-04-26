use strict;
use warnings;
use Essent;
use feature qw/say switch/;
use Switch;
use Digest::MD5;


print Digest::MD5::md5_hex "zeichen1";
print "\n";
print Digest::MD5::md5_hex "zeichen2";
