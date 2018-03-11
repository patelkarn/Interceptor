using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace Interceptor
{
    public class Input
    {
        private IntPtr context;
        private Thread callbackThread;

        /// <summary>
        /// Determines whether the driver traps no keyboard events, all events, or a range of events in-between (down only, up only...etc). Set this before loading otherwise the driver will not filter any events and no keypresses can be sent.
        /// </summary>
        public KeyboardFilterMode KeyboardFilterMode { get; set; }
        
        /// <summary>
        /// Determines whether the driver traps no events, all events, or a range of events in-between. Set this before loading otherwise the driver will not filter any events and no mouse clicks can be sent.
        /// </summary>
        public MouseFilterMode MouseFilterMode { get; set; }

        public bool IsLoaded { get; set; }

        /// <summary>
        /// Gets or sets the delay in milliseconds after each key stroke down and up. Pressing a key requires both a key stroke down and up. A delay of 0 (inadvisable) may result in no keys being apparently pressed. A delay of 20 - 40 milliseconds makes the key presses visible.
        /// </summary>
        public int KeyPressDelay { get; set; }

        /// <summary>
        /// Gets or sets the delay in milliseconds after each mouse event down and up. 'Clicking' the cursor (whether left or right) requires both a mouse event down and up. A delay of 0 (inadvisable) may result in no apparent click. A delay of 20 - 40 milliseconds makes the clicks apparent.
        /// </summary>
        public int ClickDelay { get; set; }

        public int ScrollDelay { get; set; }

        public event EventHandler<KeyPressedEventArgs> OnKeyPressed;
        public event EventHandler<MousePressedEventArgs> OnMousePressed;

        private int deviceId; /* Very important; which device the driver sends events to */

        public Input()
        {
            context = IntPtr.Zero;

            KeyboardFilterMode = KeyboardFilterMode.None;
            MouseFilterMode = MouseFilterMode.None;

            KeyPressDelay = 1;
            ClickDelay = 1;
            ScrollDelay = 15;
        }

        /*
         * Attempts to load the driver. You may get an error if the C++ library 'interception.dll' is not in the same folder as the executable and other DLLs. MouseFilterMode and KeyboardFilterMode must be set before Load() is called. Calling Load() twice has no effect if already loaded.
         */
        public bool Load()
        {
            if (IsLoaded) return false;

            context = InterceptionDriver.CreateContext();

            if (context != IntPtr.Zero)
            {
                callbackThread = new Thread(new ThreadStart(DriverCallback));
                callbackThread.Priority = ThreadPriority.Highest;
                callbackThread.IsBackground = true;
                callbackThread.Start();

                IsLoaded = true;

                return true;
            }
            else
            {
                IsLoaded = false;

                return false;
            }
        }

        /*
         * Safely unloads the driver. Calling Unload() twice has no effect.
         */
        public void Unload()
        {
            if (!IsLoaded) return;

            if (context != IntPtr.Zero)
            {
                callbackThread.Abort();
                InterceptionDriver.DestroyContext(context);
                IsLoaded = false;
            }
        }

        private void DriverCallback()
        {
            InterceptionDriver.SetFilter(context, InterceptionDriver.IsKeyboard, (Int32) KeyboardFilterMode);
            InterceptionDriver.SetFilter(context, InterceptionDriver.IsMouse, (Int32) MouseFilterMode);

            Stroke stroke = new Stroke();
            var isShiftPressed = false;

            while (InterceptionDriver.Receive(context, deviceId = InterceptionDriver.Wait(context), ref stroke, 1) > 0)
            {
                if (InterceptionDriver.IsMouse(deviceId) > 0)
                {
                    if (OnMousePressed != null)
                    {
                        var args = new MousePressedEventArgs() { X = stroke.Mouse.X, Y = stroke.Mouse.Y, State = stroke.Mouse.State, Rolling = stroke.Mouse.Rolling };
                        OnMousePressed(this, args);

                        if (args.Handled)
                        {
                            continue;
                        }
                        stroke.Mouse.X = args.X;
                        stroke.Mouse.Y = args.Y;
                        stroke.Mouse.State = args.State;
                        stroke.Mouse.Rolling = args.Rolling;
                    }
                }

                if (InterceptionDriver.IsKeyboard(deviceId) > 0)
                {
                    if (OnKeyPressed != null)
                    {
                        var hardwareIdBuffer = new byte[500];
                        var length = InterceptionDriver.GetHardwareId(context, deviceId, hardwareIdBuffer, 500);
                        var hardwareId = "";

                        var nullCount = 0;
                        for (var i = 0; i < length; i++)
                        {
                            if (hardwareIdBuffer[i] == 0)
                            {
                                nullCount++;
                                continue;
                            }

                            if (nullCount > 2)
                            {
                                hardwareId += "; ";
                            }

                            nullCount = 0;
                            hardwareId += (char)hardwareIdBuffer[i];
                        }

                        //If you need to use caps lock then be sure to have KeyboardFilterMode.All set for best reuslt
                        var isCapsLockOn = Control.IsKeyLocked(System.Windows.Forms.Keys.CapsLock);
                        if (stroke.Key.Code == Keys.LeftShift || stroke.Key.Code == Keys.RightShift)
                            isShiftPressed = stroke.Key.State == KeyState.Down;

                        var args = new KeyPressedEventArgs()
                        {
                            Key = stroke.Key.Code,
                            State = stroke.Key.State,
                            HardwareId = hardwareId,
                            KeyChar = KeyEnumToCharacter(stroke.Key.Code, isCapsLockOn, isShiftPressed)
                        };

                        OnKeyPressed?.Invoke(this, args);

                        if (args.Handled)
                        {
                            continue;
                        }
                        stroke.Key.Code = args.Key;
                        stroke.Key.State = args.State;
                    }
                }

                InterceptionDriver.Send(context, deviceId, ref stroke, 1);
            }

            Unload();
            throw new Exception("Interception.Receive() failed for an unknown reason. The driver has been unloaded.");
        }

        public void SendKey(Keys key, KeyState state)
        {
            Stroke stroke = new Stroke();
            KeyStroke keyStroke = new KeyStroke();

            keyStroke.Code = key;
            keyStroke.State = state;

            stroke.Key = keyStroke;

            InterceptionDriver.Send(context, deviceId, ref stroke, 1);

            if (KeyPressDelay > 0)
                Thread.Sleep(KeyPressDelay);
        }

        /// <summary>
        /// Warning: Do not use this overload of SendKey() for non-letter, non-number, or non-ENTER keys. It may require a special KeyState of not KeyState.Down or KeyState.Up, but instead KeyState.E0 and KeyState.E1.
        /// </summary>
        public void SendKey(Keys key)
        {
            SendKey(key, KeyState.Down);

            if (KeyPressDelay > 0)
                Thread.Sleep(KeyPressDelay);

            SendKey(key, KeyState.Up);
        }

        public void SendKeys(params Keys[] keys)
        {
            foreach (Keys key in keys)
            {
                SendKey(key);
            }
        }

        /// <summary>
        /// Warning: Only use this overload for sending letters, numbers, and symbols (those to the right of the letters on a U.S. keyboard and those obtained by pressing shift-#). Do not send special keys like Tab or Control or Enter.
        /// </summary>
        /// <param name="text"></param>
        public void SendText(string text)
        {
            foreach (char letter in text)
            {
                var tuple = CharacterToKeysEnum(letter);

                if (tuple.Item2 == true) // We need to press shift to get the next character
                    SendKey(Keys.LeftShift, KeyState.Down);

                SendKey(tuple.Item1);

                if (tuple.Item2 == true)
                    SendKey(Keys.LeftShift, KeyState.Up);
            }
        }

        /// <summary>
        /// Converts a character to a Keys enum and a 'do we need to press shift'.
        /// </summary>
        private Tuple<Keys, bool> CharacterToKeysEnum(char c)
        {
            switch (Char.ToLower(c))
            {
                case 'a':
                    return new Tuple<Keys,bool>(Keys.A, false);
                case 'b':
                    return new Tuple<Keys,bool>(Keys.B, false);
                case 'c':
                    return new Tuple<Keys,bool>(Keys.C, false);
                case 'd':
                    return new Tuple<Keys,bool>(Keys.D, false);
                case 'e':
                    return new Tuple<Keys,bool>(Keys.E, false);
                case 'f':
                    return new Tuple<Keys,bool>(Keys.F, false);
                case 'g':
                    return new Tuple<Keys,bool>(Keys.G, false);
                case 'h':
                    return new Tuple<Keys,bool>(Keys.H, false);
                case 'i':
                    return new Tuple<Keys,bool>(Keys.I, false);
                case 'j':
                    return new Tuple<Keys,bool>(Keys.J, false);
                case 'k':
                    return new Tuple<Keys,bool>(Keys.K, false);
                case 'l':
                    return new Tuple<Keys,bool>(Keys.L, false);
                case 'm':
                    return new Tuple<Keys,bool>(Keys.M, false);
                case 'n':
                    return new Tuple<Keys,bool>(Keys.N, false);
                case 'o':
                    return new Tuple<Keys,bool>(Keys.O, false);
                case 'p':
                    return new Tuple<Keys,bool>(Keys.P, false);
                case 'q':
                    return new Tuple<Keys,bool>(Keys.Q, false);
                case 'r':
                    return new Tuple<Keys,bool>(Keys.R, false);
                case 's':
                    return new Tuple<Keys,bool>(Keys.S, false);
                case 't':
                    return new Tuple<Keys,bool>(Keys.T, false);
                case 'u':
                    return new Tuple<Keys,bool>(Keys.U, false);
                case 'v':
                    return new Tuple<Keys,bool>(Keys.V, false);
                case 'w':
                    return new Tuple<Keys,bool>(Keys.W, false);
                case 'x':
                    return new Tuple<Keys,bool>(Keys.X, false);
                case 'y':
                    return new Tuple<Keys,bool>(Keys.Y, false);
                case 'z':
                    return new Tuple<Keys,bool>(Keys.Z, false);
                case '1':
                    return new Tuple<Keys,bool>(Keys.One, false);
                case '2':
                    return new Tuple<Keys,bool>(Keys.Two, false);
                case '3':
                    return new Tuple<Keys,bool>(Keys.Three, false);
                case '4':
                    return new Tuple<Keys,bool>(Keys.Four, false);
                case '5':
                    return new Tuple<Keys,bool>(Keys.Five, false);
                case '6':
                    return new Tuple<Keys,bool>(Keys.Six, false);
                case '7':
                    return new Tuple<Keys,bool>(Keys.Seven, false);
                case '8':
                    return new Tuple<Keys,bool>(Keys.Eight, false);
                case '9':
                    return new Tuple<Keys,bool>(Keys.Nine, false);
                case '0':
                    return new Tuple<Keys,bool>(Keys.Zero, false);
                case '-':
                    return new Tuple<Keys,bool>(Keys.DashUnderscore, false);
                case '+':
                    return new Tuple<Keys,bool>(Keys.PlusEquals, false);
                case '[':
                    return new Tuple<Keys,bool>(Keys.OpenBracketBrace, false);
                case ']':
                    return new Tuple<Keys,bool>(Keys.CloseBracketBrace, false);
                case ';':
                    return new Tuple<Keys,bool>(Keys.SemicolonColon, false);
                case '\'':
                    return new Tuple<Keys,bool>(Keys.SingleDoubleQuote, false);
                case ',':
                    return new Tuple<Keys,bool>(Keys.CommaLeftArrow, false);
                case '.':
                    return new Tuple<Keys,bool>(Keys.PeriodRightArrow, false);
                case '/':
                    return new Tuple<Keys,bool>(Keys.ForwardSlashQuestionMark, false);
                case '{':
                    return new Tuple<Keys,bool>(Keys.OpenBracketBrace, true);
                case '}':
                    return new Tuple<Keys,bool>(Keys.CloseBracketBrace, true);
                case ':':
                    return new Tuple<Keys,bool>(Keys.SemicolonColon, true);
                case '\"':
                    return new Tuple<Keys,bool>(Keys.SingleDoubleQuote, true);
                case '<':
                    return new Tuple<Keys,bool>(Keys.CommaLeftArrow, true);
                case '>':
                    return new Tuple<Keys,bool>(Keys.PeriodRightArrow, true);
                case '?':
                    return new Tuple<Keys,bool>(Keys.ForwardSlashQuestionMark, true);
                case '\\':
                    return new Tuple<Keys,bool>(Keys.BackslashPipe, false);
                case '|':
                    return new Tuple<Keys,bool>(Keys.BackslashPipe, true);
                case '`':
                    return new Tuple<Keys,bool>(Keys.Tilde, false);
                case '~':
                    return new Tuple<Keys,bool>(Keys.Tilde, true);
                case '!':
                    return new Tuple<Keys,bool>(Keys.One, true);
                case '@':
                    return new Tuple<Keys,bool>(Keys.Two, true);
                case '#':
                    return new Tuple<Keys,bool>(Keys.Three, true);
                case '$':
                    return new Tuple<Keys,bool>(Keys.Four, true);
                case '%':
                    return new Tuple<Keys,bool>(Keys.Five, true);
                case '^':
                    return new Tuple<Keys,bool>(Keys.Six, true);
                case '&':
                    return new Tuple<Keys,bool>(Keys.Seven, true);
                case '*':
                    return new Tuple<Keys,bool>(Keys.Eight, true);
                case '(':
                    return new Tuple<Keys,bool>(Keys.Nine, true);
                case ')':
                    return new Tuple<Keys, bool>(Keys.Zero, true);
                case ' ':
                    return new Tuple<Keys, bool>(Keys.Space, true);
                default:
                    return new Tuple<Keys, bool>(Keys.ForwardSlashQuestionMark, true);
            }
        }

        /// <summary>
        /// Converts a key enum to char.
        /// </summary>
        private char KeyEnumToCharacter(Keys keys, bool isCapsLockOn, bool isShiftPressed)
        {
            var capitalizeCharacter = (isCapsLockOn && !isShiftPressed) || (!isCapsLockOn && isShiftPressed);
            switch (keys)
            {
                case Keys.A:
                    return capitalizeCharacter ? 'A' : 'a';
                case Keys.B:
                    return capitalizeCharacter ? 'B' : 'b';
                case Keys.C:
                    return capitalizeCharacter ? 'C' : 'c';
                case Keys.D:
                    return capitalizeCharacter ? 'D' : 'd';
                case Keys.E:
                    return capitalizeCharacter ? 'E' : 'e';
                case Keys.F:
                    return capitalizeCharacter ? 'F' : 'f';
                case Keys.G:
                    return capitalizeCharacter ? 'G' : 'g';
                case Keys.H:
                    return capitalizeCharacter ? 'H' : 'h';
                case Keys.I:
                    return capitalizeCharacter ? 'I' : 'i';
                case Keys.J:
                    return capitalizeCharacter ? 'J' : 'j';
                case Keys.K:
                    return capitalizeCharacter ? 'K' : 'k';
                case Keys.L:
                    return capitalizeCharacter ? 'L' : 'l';
                case Keys.M:
                    return capitalizeCharacter ? 'M' : 'm';
                case Keys.N:
                    return capitalizeCharacter ? 'N' : 'n';
                case Keys.O:
                    return capitalizeCharacter ? 'O' : 'o';
                case Keys.P:
                    return capitalizeCharacter ? 'P' : 'p';
                case Keys.Q:
                    return capitalizeCharacter ? 'Q' : 'q';
                case Keys.R:
                    return capitalizeCharacter ? 'R' : 'r';
                case Keys.S:
                    return capitalizeCharacter ? 'S' : 's';
                case Keys.T:
                    return capitalizeCharacter ? 'T' : 't';
                case Keys.U:
                    return capitalizeCharacter ? 'U' : 'u';
                case Keys.V:
                    return capitalizeCharacter ? 'V' : 'v';
                case Keys.W:
                    return capitalizeCharacter ? 'W' : 'w';
                case Keys.X:
                    return capitalizeCharacter ? 'X' : 'x';
                case Keys.Y:
                    return capitalizeCharacter ? 'Y' : 'y';
                case Keys.Z:
                    return capitalizeCharacter ? 'Z' : 'z';
                case Keys.One:
                    return isShiftPressed ? '!' : '1';
                case Keys.Two:
                    return isShiftPressed ? '@' : '2';
                case Keys.Three:
                    return isShiftPressed ? '#' : '3';
                case Keys.Four:
                    return isShiftPressed ? '$' : '4';
                case Keys.Five:
                    return isShiftPressed ? '%' : '5';
                case Keys.Six:
                    return isShiftPressed ? '^' : '6';
                case Keys.Seven:
                    return isShiftPressed ? '&' : '7';
                case Keys.Eight:
                    return isShiftPressed ? '*' : '8';
                case Keys.Nine:
                    return isShiftPressed ? '(' : '9';
                case Keys.Zero:
                    return isShiftPressed ? ')' : '0';
                case Keys.DashUnderscore:
                    return isShiftPressed ? '_' : '-';
                case Keys.PlusEquals:
                    return isShiftPressed ? '+' : '=';
                case Keys.OpenBracketBrace:
                    return isShiftPressed ? '{' : '[';
                case Keys.CloseBracketBrace:
                    return isShiftPressed ? '}' : ']';
                case Keys.SemicolonColon:
                    return isShiftPressed ? ':' : ';';
                case Keys.SingleDoubleQuote:
                    return isShiftPressed ? '"' : '\'';
                case Keys.CommaLeftArrow:
                    return isShiftPressed ? '<' : ',';
                case Keys.PeriodRightArrow:
                    return isShiftPressed ? '>' : '.';
                case Keys.ForwardSlashQuestionMark:
                    return isShiftPressed ? '?' : '/';
                case Keys.BackslashPipe:
                    return isShiftPressed ? '|' : '\\';
                case Keys.Tilde:
                    return isShiftPressed ? '~' : '`';
                case Keys.Space:
                    return ' ';
                case Keys.Enter:
                    return (char)13;
                case Keys.Numpad0:
                    return '0';
                case Keys.Numpad1:
                    return '1';
                case Keys.Numpad2:
                    return '2';
                case Keys.Numpad3:
                    return '3';
                case Keys.Numpad4:
                    return '4';
                case Keys.Numpad5:
                    return '5';
                case Keys.Numpad6:
                    return '6';
                case Keys.Numpad7:
                    return '7';
                case Keys.Numpad8:
                    return '8';
                case Keys.Numpad9:
                    return '9';
                case Keys.NumpadAsterisk:
                    return '*';
                case Keys.NumpadMinus:
                    return '-';
                case Keys.NumpadPlus:
                    return '+';
                default:
                    return (char)0;
            }
        }

        public void SendMouseEvent(MouseState state)
        {
            Stroke stroke = new Stroke();
            MouseStroke mouseStroke = new MouseStroke();

            mouseStroke.State = state;

            if (state == MouseState.ScrollUp)
            {
                mouseStroke.Rolling = 120;
            }
            else if (state == MouseState.ScrollDown)
            {
                mouseStroke.Rolling = -120;
            }

            stroke.Mouse = mouseStroke;

            InterceptionDriver.Send(context, 12, ref stroke, 1);
        }

        public void SendLeftClick()
        {
            SendMouseEvent(MouseState.LeftDown);
            Thread.Sleep(ClickDelay);
            SendMouseEvent(MouseState.LeftUp);
        }

        public void SendRightClick()
        {
            SendMouseEvent(MouseState.RightDown);
            Thread.Sleep(ClickDelay);
            SendMouseEvent(MouseState.RightUp);
        }

        public void ScrollMouse(ScrollDirection direction)
        {
            switch (direction)
            { 
                case ScrollDirection.Down:
                    SendMouseEvent(MouseState.ScrollDown);
                    break;
                case ScrollDirection.Up:
                    SendMouseEvent(MouseState.ScrollUp);
                    break;
            }
        }

        /// <summary>
        /// Warning: This function, if using the driver, does not function reliably and often moves the mouse in unpredictable vectors. An alternate version uses the standard Win32 API to get the current cursor's position, calculates the desired destination's offset, and uses the Win32 API to set the cursor to the new position.
        /// </summary>
        public void MoveMouseBy(int deltaX, int deltaY, bool useDriver = false)
        {
            if (useDriver)
            {
                Stroke stroke = new Stroke();
                MouseStroke mouseStroke = new MouseStroke();

                mouseStroke.X = deltaX;
                mouseStroke.Y = deltaY;

                stroke.Mouse = mouseStroke;
                stroke.Mouse.Flags = MouseFlags.MoveRelative;

                InterceptionDriver.Send(context, 12, ref stroke, 1);
            }
            else
            {
                var currentPos = Cursor.Position;
                Cursor.Position = new Point(currentPos.X + deltaX, currentPos.Y - deltaY); // Coordinate system for y: 0 begins at top, and bottom of screen has the largest number
            }
        }

        /// <summary>
        /// Warning: This function, if using the driver, does not function reliably and often moves the mouse in unpredictable vectors. An alternate version uses the standard Win32 API to set the cursor's position and does not use the driver.
        /// </summary>
        public void MoveMouseTo(int x, int y, bool useDriver = false)
        {
            if (useDriver)
            {
                Stroke stroke = new Stroke();
                MouseStroke mouseStroke = new MouseStroke();

                mouseStroke.X = x;
                mouseStroke.Y = y;

                stroke.Mouse = mouseStroke;
                stroke.Mouse.Flags = MouseFlags.MoveAbsolute;

                InterceptionDriver.Send(context, 12, ref stroke, 1);
            }
            {
                Cursor.Position = new Point(x, y);
            }
        }
    }
}
 