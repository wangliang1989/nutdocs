#!/usr/bin/env perl
use strict;
use warnings;

my $cmd = '/Applications/LibreOffice.app/Contents/MacOS/soffice --headless --convert-to csv';

chdir "xlsx" or die;
foreach my $class (glob "*") {
    chdir $class or next;
    foreach my $test (glob "*") {
        chdir $test or next;
        system "$cmd $_" foreach (glob "*.xlsx");
        chdir ".." or die;
    }
    chdir '..' or die;
}
