use warnings;
use strict;

use Win32::OLE;

my $wsh=new Win32::OLE("WScript.Shell") or die;

print $wsh -> SpecialFolders("Fonts");
