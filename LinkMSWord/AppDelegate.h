//
//  AppDelegate.h
//  LinkMSWord
//
//  Created by Rameshkumar on 18/11/13.
//  Copyright (c) 2013 RYG. All rights reserved.
//

#import <Cocoa/Cocoa.h>
#import "MicrosoftWord.h"

@interface AppDelegate : NSObject <NSApplicationDelegate, SBApplicationDelegate>

@property (assign) IBOutlet NSWindow *window;
@property (nonatomic, strong) MicrosoftWordApplication *wordApp;
- (IBAction)openWordDoc:(id)sender;

@end
