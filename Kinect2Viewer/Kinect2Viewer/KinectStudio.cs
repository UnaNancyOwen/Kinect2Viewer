#region
// * KinectStudio Class
//     This class provide the controls (play, pause, resume, stop) of clip that recorded by Kinect Studio.
//
// * License
//    Copyright (c) 2016 Tsukasa SUGIURA
//    This source code is licensed under the MIT license.
//
// * How to setup Microsoft.Kinect.KinectStudio
//     1. Add This Item to a Project.
//     2. Configure a Project to Target 64-Bit(x64) Platforms. (because "Microsoft.Kinect.Tools" support only x64.)
//     3. Add Reference Microsoft.Kinect.Tools.dll to a Project.
//     4. Add Commands "xcopy "$(KINECTSDK20_DIR)Tools\KinectStudio\KStudioService.dll" "$(TargetDir)" /S /R /Y /I" to Post-Build Event Command Line.
//     5. Add "using Microsoft.Kinect.KinectStudio;" to Your Source Code.
#endregion

using System;
using System.IO;
using System.Threading;
using Microsoft.Kinect.Tools;

namespace Microsoft.Kinect.KinectStudio
{
    /// <summary>
    /// This class provide the controls (play, pause, resume, stop) of clip that recorded by Kinect Studio.
    /// </summary>
    public class KinectStudio
    {
        private KStudioClient client;
        private Thread thread;
        private string path;
        private uint loop;
        private bool isPause;

        /// <summary>
        /// Constructor
        /// </summary>
        public KinectStudio()
        {
            client = KStudio.CreateClient();
            thread = null;
            path = "";
            loop = 0;
            isPause = false;
        }

        /// <summary>
        /// Constructor with Settings
        /// </summary>
        /// <param name="path">Absolute path to clip.</param>
        /// <param name="loop">Number of times to loop.</param>
        public KinectStudio(string path, uint loop = 0)
        {
            if (!Path.IsPathRooted(path))
            {
                throw new ArgumentException("Need Enter Absolute Path to Clip", "path");
            }

            client = KStudio.CreateClient();
            thread = null;
            this.path = path;
            this.loop = loop;
            isPause = false;
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~KinectStudio()
        {
            client.Dispose();
            client = null;
        }

        /// <summary>
        /// Clip
        /// </summary>
        /// <param name="path">Absolute path to clip.</param>
        /// <param name="loop">Number of times to loop.</param>
        public void Clip(string path, uint loop = 0)
        {
            if (!Path.IsPathRooted(path))
            {
                throw new ArgumentException("Need Enter Absolute Path to Clip", "path");
            }

            if (!this.path.Equals(path))
            {
                Stop();
            }

            this.path = path;
            this.loop = loop;
        }

        /// <summary>
        /// Play
        /// </summary>
        public void Play()
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            if (thread != null)
            {
                if (!thread.ThreadState.Equals(ThreadState.Stopped))
                {
                    Resume();
                    return;
                }
            }

            isPause = false;

            thread = new Thread(new ThreadStart(Run));
            thread.Start();
        }

        /// <summary>
        /// Pause
        /// </summary>
        public void Pause()
        {
            if (thread != null)
            {
                if (!thread.ThreadState.Equals(ThreadState.Stopped))
                {
                    isPause = true;
                }
            }
        }

        /// <summary>
        /// Resume
        /// </summary>
        public void Resume()
        {
            if (thread != null)
            {
                if (!thread.ThreadState.Equals(ThreadState.Stopped))
                {
                    isPause = false;
                }
            }
        }

        /// <summary>
        /// Stop
        /// </summary>
        public void Stop()
        {
            if (thread != null)
            {
                client.DisconnectFromService();

                thread.Abort();
                thread.Join();
                thread = null;
            }
        }

        /// <summary>
        /// Run
        /// </summary>
        private void Run()
        {
            client.ConnectToService();

            using (KStudioPlayback play = client.CreatePlayback(path))
            {
                play.EndBehavior = KStudioPlaybackEndBehavior.Stop;
                play.Mode = KStudioPlaybackMode.TimingEnabled;
                play.LoopCount = loop;
                play.Start();

                while (play.State.Equals(KStudioPlaybackState.Playing) || play.State.Equals(KStudioPlaybackState.Paused))
                {
                    Thread.Sleep(33);

                    if (isPause && !play.State.Equals(KStudioPlaybackState.Paused))
                    {
                        play.Pause();
                    }

                    if (!isPause && play.State.Equals(KStudioPlaybackState.Paused))
                    {
                        play.Resume();
                    }

                    if (play.State.Equals(KStudioPlaybackState.Error))
                    {
                        throw new InvalidOperationException("KStudioPlayback Error");
                    }
                }

                play.Stop();
            }

            client.DisconnectFromService();
        }
    }
}
