//------------------------------------------------------------------------------
// <copyright file="TabGroupJump.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE80;
using EnvDTE;
using System.Collections.Generic;

namespace TabGroupJumperVSIX
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TabGroupJump
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandIdJumpLeft = 0x0100;
        public const int CommandIdJumpRight = 0x0101;
        public const int CommandIdJumpUp = 0x0102;
        public const int CommandIdJumpDown = 0x0103;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("67e760eb-74c3-46f2-9b85-c7af2d351428");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="TabGroupJump"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private TabGroupJump(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                // Add the jump commands to the handler
                commandService.AddCommand(
                    new MenuCommand(this.MenuItemCallback, new CommandID(CommandSet, CommandIdJumpLeft)));
                commandService.AddCommand(
                    new MenuCommand(this.MenuItemCallback, new CommandID(CommandSet, CommandIdJumpRight)));
                commandService.AddCommand(
                    new MenuCommand(this.MenuItemCallback, new CommandID(CommandSet, CommandIdJumpUp)));
                commandService.AddCommand(
                    new MenuCommand(this.MenuItemCallback, new CommandID(CommandSet, CommandIdJumpDown)));
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static TabGroupJump Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new TabGroupJump(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            DTE2 dte = (DTE2)this.ServiceProvider.GetService(typeof(DTE));
            int commandId = ((MenuCommand)sender).CommandID.ID;
            bool horizontal = commandId == CommandIdJumpLeft || commandId == CommandIdJumpRight;

            // Documents with a "left" or "top" value > 0 are the focused ones in each group, 
            // so we only need to collect those
            List<EnvDTE.Window> topLevel = new List<EnvDTE.Window>();
            foreach (EnvDTE.Window w in dte.Windows)
            {
                if (w.Kind == "Document" && (horizontal && w.Left > 0 || w.Top > 0))
                    topLevel.Add(w);
            }

            if (topLevel.Count == 0)
                return;

            if (horizontal)
                topLevel.Sort((a, b) => a.Left < b.Left ? -1 : 1);
            else
                topLevel.Sort((a, b) => a.Top < b.Top ? -1 : 1);

            // find the index of the active document
            var activeDoc = dte.ActiveDocument;
            int activeIdx = 0;
            for (int i = 0; i < topLevel.Count; ++i)
            {
                if (topLevel[i].Document == activeDoc)
                {
                    activeIdx = i;
                    break;
                }
            }

            // set the new active document
            if (horizontal)
            {
                activeIdx += commandId == CommandIdJumpLeft ? -1 : 1;
            }
            else
            {
                activeIdx += commandId == CommandIdJumpUp ? -1 : 1;
            }

            activeIdx = (activeIdx < 0 ? activeIdx + topLevel.Count : activeIdx) % topLevel.Count;
            topLevel[activeIdx].Activate();
        }
    }
}
