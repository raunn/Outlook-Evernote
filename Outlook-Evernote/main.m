//
//  main.m
//  Outlook-Evernote
//
//  Created by Raun Nohavitza on 12/23/14.
//  Copyright (c) 2014 Raun Nohavitza. All rights reserved.
//

#import <Cocoa/Cocoa.h>
#import <AppleScriptObjC/AppleScriptObjC.h>

int main(int argc, const char * argv[]) {
    [[NSBundle mainBundle] loadAppleScriptObjectiveCScripts];
    return NSApplicationMain(argc, argv);
}
