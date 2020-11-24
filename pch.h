// pch.h: This is a precompiled header file.
// Files listed below are compiled only once, improving build performance for future builds.
// This also affects IntelliSense performance, including code completion and many code browsing features.
// However, files listed here are ALL re-compiled if any one of them is updated between builds.
// Do not add files here that you will be updating frequently as this negates the performance advantage.

#ifndef PCH_H
#define PCH_H

// add headers that you want to pre-compile here
#include "framework.h"

#pragma warning (disable : 4146)
#import "C:\Program Files\Common Files\DESIGNER\MSADDNDR.OLB" raw_interfaces_only, raw_native_types, no_namespace, named_guids, auto_search

// MSADDNDR.OLB"
//#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4"\
//auto_rename auto_search raw_interfaces_only rename_namespace("AddinDesign")

// Office type library (i.e., mso.dll).
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"\
    auto_rename auto_search raw_interfaces_only rename_namespace("Office")

#import "libid:00020905-0000-0000-C000-000000000046"\
    auto_rename auto_search raw_interfaces_only rename_namespace("Word")

//using namespace AddinDesign;
using namespace Office;
using namespace Word;
#endif //PCH_H
