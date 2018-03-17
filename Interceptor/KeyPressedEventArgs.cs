﻿using System;

namespace Interceptor
{
    public class KeyPressedEventArgs : EventArgs
    {
        public Keys Key { get; set; }
        public KeyState State { get; set; }
        public bool Handled { get; set; }
        public string HardwareId { get; set; }
        public int AsciiKeyCode { get; set; }
    }
}
