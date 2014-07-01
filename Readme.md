#  org-outlook.el
 Matthew L. Fidler
## Library Information
 _org-outlook.el_ --- Outlook org

- __Filename__ --  org-outlook.el
Description: 
- __Author__ --  Matthew L. Fidler
Maintainer:
- __Created__ --  Mon May 10 09:44:59 2010 (-0500)
- __Version__ --  0.11
- __Last-Updated__ --  Tue May 29 22:21:06 2012 (-0500)
- __By__ --  Matthew L. Fidler
- __Update #__ --  166
- __URL__ --  https:__github.com_mlf176f2_org-outlook.el
- __Keywords__ --  Org-outlook 
Compatibility:

## Introduction:
Org mode lets you organize your tasks. However, sometimes you may wish
to integrate org-mode with outlook since your company forces you to
use Microsoft Outlook.  org-outlook.el allows: 

- Creating Tasks from outlook items:
  - org-outlook-task. All selected items in outlook will be added to a
    task-list at current point. This version requires org-protocol and   
    org-protocol.vbs.  The org-protocol.vbs has can be generated with
    the interactive function `org-outlook-create-vbs`.

  - If your organization has blocked all macro access OR you want to
    have an action for a saved `.msg` email, org-outlook also adds
    drag and drop support allowing `.msg` files to become org tasks.
    This is enabled by default, but can be disabled by
    `org-outlook-no-dnd`

  - With blocked emails, you may wish to delete the emails in a folder
    after the task is completed.  This can be accomplished with
    `org-protocol-delete-msgs`.  If you use it frequently, you may
    wish to bind it to a key, like


  (define-key org-mode-map (kbd "C-c d") 'org-protocol-delete-msgs)



- Open Outlook Links in org-mode

  - Requires org-outlook-location to be customized when using Outlook
    2007 (this way you don’t have to edit the registry).

This is based loosely on:
http:__superuser.com_questions_71786/can-i-create-a-link-to-a-specific-email-message-in-outlook


Note that you may also add tasks using visual basic directly. The script below performs the following actions:

   - Move email to Personal Folders under folder "@ActionTasks" (changes GUID)
   - Create a org-mode task under heading "* Tasks" for the file `f:\Documents\org\gtd.org`
   - Note by replacing "@ActionTasks", "* Tasks" and
     `f:\Documents\org\gtd.org` you can modify this script to your
     personal needs.

The visual basic script for outlook can be created by calling `M-x org-outlook-create-vbs`

## History

1-Jul-2014    Matthew L. Fidler  
   Last-Updated: Tue May 29 22:21:06 2012 (-0500) #166 (Matthew L. Fidler)
   Add delete msg files support
- __24-Jun-2014__ --   Bugfix for Drag and Drop Support (Matthew L. Fidler)
- __24-Jun-2014__ --   Add Drag and drop support for tasks (Matthew L. Fidler)
- __12-Dec-2012__ --   Updated Visual Basic Script to be more robust, and have more options. (Matthew L. Fidler)
- __07-Dec-2012__ --   Should fix Issue #1. Also added org-outlook-create-vbs to create the VBS code based on a user's setup. (Matthew L. Fidler)
- __26-May-2012__ --   Added (require 'cl), Thanks Robert Pluim (Matthew L. Fidler)
- __21-Feb-2012__ --   Bug fix for opening files. (Matthew L. Fidler)
- __21-Feb-2012__ --   Bug fix. (Matthew L. Fidler)
- __13-Dec-2011__ --   Added more autoload cookies. (Matthew L. Fidler)
- __08-Apr-2011__ --   Added some autoload cookies. (US041375)
- __15-Feb-2011__ --   Changed outlook-org to org-outlook.el (Matthew L. Fidler)
- __11-Jan-2011__ --   Finalized interface with org-protocol (Matthew L. Fidler)
- __05-Jan-2011__ --   Removed outlook copy. I only use from outlook now.  (Matthew L. Fidler)
