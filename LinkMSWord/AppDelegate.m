//
//  AppDelegate.m
//  LinkMSWord
//
//  Created by Rameshkumar on 18/11/13.
//  Copyright (c) 2013 RYG. All rights reserved.
//

#import "AppDelegate.h"

@implementation AppDelegate

- (void)applicationDidFinishLaunching:(NSNotification *)aNotification
{
    // Insert code here to initialize your application
}

- (IBAction)openWordDoc:(id)sender{
    
    NSOpenPanel *openDlg = [NSOpenPanel openPanel];
    [openDlg setCanChooseFiles:YES];
    [openDlg setCanChooseDirectories:NO];
    [openDlg setAllowsMultipleSelection:NO];
    [openDlg beginSheetModalForWindow:self.window completionHandler:^(NSInteger result){
        if (result == 1) {
            NSMutableString *fileUrl = [[NSMutableString alloc] initWithString:[[openDlg URL]description]] ;
            NSRange range = [fileUrl rangeOfString:@"file://"];
            NSString *filePath = [fileUrl stringByReplacingCharactersInRange:range withString:@"file://localhost"];
            [self openWordDocument:filePath];
        }
    }];
    
}
- (void) openWordDocument:(NSString *)FilePath{
    
    self.wordApp = [SBApplication applicationWithBundleIdentifier:@"com.microsoft.word"];
    self.wordApp.delegate = self;
    [self.wordApp activate];
    
    if ([self.wordApp isRunning]) {
        MicrosoftWordDocument *activeDocument = [self.wordApp activeDocument];
        [activeDocument openFileName:FilePath
                  confirmConversions:NO
                            readOnly:NO
                    addToRecentFiles:NO
                    passwordDocument:NO
                    passwordTemplate:NO
                              Revert:NO
                       writePassword:NO
               writePasswordTemplate:NO
                       fileConverter:MicrosoftWordE162OpenFormatDocument97];
    }
    
}

-(id)eventDidFail:(const AppleEvent *)event withError:(NSError *)error{
    NSLog(@"Cannot open %@", error);
    return nil;
}


@end
