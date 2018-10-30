﻿using System;
using System.Linq;
using System.Windows.Input;

namespace TemporaryProjects
{
    internal sealed class CombinedCommand : ICommand
    {
        private readonly ICommand[] commands;

        public CombinedCommand(params ICommand[] commands) => this.commands = commands;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter) => commands.All(c => c.CanExecute(parameter));

        public void Execute(object parameter)
        {
            foreach (var command in commands)
                command.Execute(parameter);
        }
    }
}
