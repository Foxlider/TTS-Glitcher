using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Media;
using System.Speech.Synthesis;
using System.Threading.Tasks;

namespace TTSGlitcher
{
    class Program
    {
        private static SpeechSynthesizer _synth;
        private static List<VoiceInfo>   _voices   = new List<VoiceInfo>();
        const          int               BlockSize = 22050;
        private static Random            _r;

        private static async Task Main(string[] args)
        {
            #region vars
            _r = new Random();
            _synth = new SpeechSynthesizer();
            #endregion

            #region args parser
            if (args.Length < 3)
            {
                Console.WriteLine("No arguments sent. Please provide the arguments for the application to run\n");
                Console.WriteLine("- string     Path of the output .wav file\n");
                Console.WriteLine("- string     Language for the bot (en-US/fr-FR)\n");
                Console.WriteLine("- string[]   Text to say\n\nPress any key to exit...");
                Console.ReadKey();
                Environment.Exit(-1);
            }
            #endregion

            #region Voice init
            // Output information about all of the installed voices.
            _voices = _synth.GetInstalledVoices()
                            .Where(v =>
                                       (v.VoiceInfo.Id == "MSTTS_V110_enUS_EvaM")
                                    || (v.VoiceInfo.Id == "MSTTS_V110_frFR_NathalieM"))
                            .Select(v => v.VoiceInfo)
                            .ToList();
            if (_voices == null || _voices.Count < 1)
            {
                Console.WriteLine("Voices were not recognized. Please install the required TTS voices to proceed.\nPress any key to exit...");
                Console.ReadKey();
                Environment.Exit(-2);
            }
            #endregion

            #region TTS stream creation
            var said = string.Join(" ", args.Skip(2));
            Console.WriteLine("Output text to file > "+ args[0]);

            var stream = TtsToStream(said, args[1]);
            Console.WriteLine("Done.");
            #endregion

            #region Extract Stream's header
            stream.Seek(0, SeekOrigin.Begin);
            var ms = new MemoryStream();
            List<(string, byte[], object)> headerVals;
            headerVals = GetHeader(stream); //Get all WAV File header values
            byte[] header = (byte[]) GetValueFromHeader(headerVals, "fullHeader", false); //Extract the full header
            ms.Position = header.Length;        // |Set the streams position
            stream.Position = header.Length;    // |

            var totalbytes = stream.Length;
            int streamChunkSize = (int)GetValueFromHeader(headerVals, "chunkSize");
            int streamDataSize = (int) GetValueFromHeader(headerVals, "dataChunkSize");
            var bitsPerSample = (short) GetValueFromHeader(headerVals, "bitsPerSample");
            #endregion

            #region Glitching the stream
            MemoryStream msLQ = new MemoryStream();
            while (true) //Low Quality pass
            {
                var buffer    = new byte[BlockSize];
                var byteCount = await stream.ReadAsync(buffer, 0, BlockSize);
                if (byteCount == 0) { break; }
                var strm = new MemoryStream();
                bool passing = true;
                for (int i = 0; i < buffer.Length; i += 2)
                {
                    if (passing)
                    {
                        strm.WriteByte(buffer[i]);
                        strm.WriteByte(buffer[i + 1]);
                        strm.WriteByte(buffer[i]);
                        strm.WriteByte(buffer[i + 1]);
                    }
                    passing = !passing;
                }
                await msLQ.WriteAsync(strm.ToArray(), 0, (int)strm.Length);
            }
            ms.Position = header.Length;
            msLQ.Position = header.Length;
            while (true) //Copy stream to memorystream (to glitch it eventually)
            {
                var       buffer    = new byte[BlockSize];
                var       byteCount = await msLQ.ReadAsync(buffer, 0, BlockSize);
                if (byteCount == 0) { break; }
                await Glitcher(ms, buffer, byteCount, bitsPerSample);
                //await ms.WriteAsync(buffer, 0, byteCount);
            }
            #endregion

            #region recreate Stream's header
            long newTotalBytes = ms.Length;
            int  newChunkSize  = (int) (newTotalBytes - (totalbytes - streamChunkSize));
            int  newDataSize   = (int) (newTotalBytes - (totalbytes - streamDataSize));

            List<(string, byte[], object)> newHeader = headerVals;
            headerVals[GetIndexFromHeader(headerVals, "chunkSize")] = ("chunkSize", BitConverter.GetBytes(newChunkSize), newChunkSize);
            headerVals[GetIndexFromHeader(headerVals, "dataChunkSize")] = ("dataChunkSize", BitConverter.GetBytes(newDataSize), newDataSize);
            byte[] newBytes = SetHeader(newHeader).Item1;
            
            ms.Seek(0, SeekOrigin.Begin);
            ms.Write(newBytes, 0, newBytes.Length); //Rewrite the stream's header
            GetHeader(ms);
            #endregion

            #region Write stream to file
            ms.Seek(0, SeekOrigin.Begin);
            Console.WriteLine("Finished copying stream");
            FileStream fs = new FileStream(args[0], FileMode.OpenOrCreate); //Write memorystream to file
            ms.WriteTo(fs);
            #endregion

            #region Play down stream
            ms.Seek(0, SeekOrigin.Begin);
            var player = new SoundPlayer(ms); //Play the audio stream
            player.PlaySync();
            #endregion 

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        /// <summary>
        /// Can glitch up a stream of data
        /// </summary>
        /// <param name="ms">MemoryStream of data to write on</param>
        /// <param name="buffer">Buffer of data to glitch up</param>
        /// <param name="byteCount">bytecount of data buffer</param>
        /// <returns></returns>
        private static async Task Glitcher(MemoryStream ms, byte[] buffer, int byteCount, int bps)
        {
            var strm = new MemoryStream();
            switch (_r.Next(20))
            {
                #region G1
                case 1:
                case 2:
                case 3:
                case 4:
                    
                    int fragLen = buffer.Length / 8;
                    byte[] frag = buffer.Take(fragLen).ToArray();
                    //byte[] frag = buffer.Take(2756).ToArray();
                    var loops = _r.Next(3, 10);
                    for (var i = 0; i <= loops; i++)
                    {
                        strm.Write(frag, 0, frag.Length);
                        //Console.WriteLine($"Looped and created  a buffer {strm.Length} ({frag.Length}) == {strm.ToArray().Length}");
                    }
                    Console.WriteLine($"G1 ({loops} loops)");
                    byte[] endFrag = buffer.ToList().GetRange(fragLen, buffer.Length - fragLen).ToArray();
                    strm.Write(endFrag, 0, endFrag.Length);
                    break;
                #endregion

                //#region G2
                //case 4:
                //    Console.WriteLine("G2");
                //    for (int i = 0; i < buffer.Length; i += 2)
                //    {
                //        strm.WriteByte(buffer[i]);
                //        strm.WriteByte(buffer[i + 1]);
                //        strm.WriteByte(buffer[i]);
                //        strm.WriteByte(buffer[i + 1]);
                //    }
                //    byteCount = (int)strm.Length;
                //    break;
                //#endregion

                #region G4
                case 5:
                case 6:
                    Console.WriteLine("G4");
                    for (int i = 0; i < buffer.Length; i += 2)
                    {
                        switch (_r.Next(3))
                        {
                            case 1:
                                strm.WriteByte(buffer[i]);
                                strm.WriteByte(buffer[i + 1]);
                                strm.WriteByte(buffer[i]);
                                strm.WriteByte(buffer[i + 1]);
                                break;
                            case 2:
                                strm.WriteByte(buffer[i]);
                                strm.WriteByte(buffer[i + 1]);
                                strm.WriteByte(buffer[i]);
                                strm.WriteByte(buffer[i + 1]);
                                strm.WriteByte(buffer[i]);
                                strm.WriteByte(buffer[i + 1]);
                                break;
                                //case 0 or 3 just skips a value because lulz
                        }
                    }
                    break;
                #endregion

                #region G5
                case 7:
                case 8:
                case 9:
                case 10:
                    Console.WriteLine("G5");
                    bool passing = true;
                    for (int i = 0; i < buffer.Length; i += 2)
                    {
                        if (passing)
                        {
                            strm.WriteByte(buffer[i]);
                            strm.WriteByte(buffer[i + 1]);
                            strm.WriteByte(buffer[i]);
                            strm.WriteByte(buffer[i + 1]);
                            strm.WriteByte(buffer[i]);
                            strm.WriteByte(buffer[i + 1]);
                            strm.WriteByte(buffer[i]);
                            strm.WriteByte(buffer[i + 1]);
                        }
                        else { i += 4; }
                        passing = !passing;
                    }
                    break;
                #endregion

                default:
                    strm.Write(buffer, 0, byteCount);
                    break;

            }
            await ms.WriteAsync(strm.ToArray(), 0, (int)strm.Length);
        }

        /// <summary>
        /// Transforms a string text to a TTS stream with the selected language 
        /// </summary>
        /// <param name="said">Text to send to the TTS System</param>
        /// <param name="culture">Only 'en-US' and 'fr-FR' values are supported for now</param>
        /// <returns></returns>
        private static Stream TtsToStream(string said, string culture = "en-US")
        {
            if (_synth == null) { throw new InvalidOperationException("No speech synthetizer have been initiated."); }
            if (culture != "en-US" && culture != "fr-FR")
            { throw new InvalidOperationException("Only 'en-US' and 'fr-FR' cultures are supported right now."); }
            var stream = new MemoryStream();
            _synth.SetOutputToWaveStream(stream);
            PromptBuilder builder = new PromptBuilder(new System.Globalization.CultureInfo(culture));
            builder.StartParagraph();
            builder.AppendText(said);
            builder.EndParagraph();

            try
            { _synth.SelectVoice(GetVoice(culture)); }
            catch
            { throw new InvalidOperationException($"Could not select a voice with culture {culture}."); }

            _synth.Speak(said);

            return stream;
        }

        /// <summary>
        /// Get a voice's name depending on the culture
        /// </summary>
        /// <param name="culture">Voice's culture</param>
        /// <returns></returns>
        private static string GetVoice(string culture)
        { return _voices.Where(v => v.Culture.ToString() == culture).Select(v => v.Name).FirstOrDefault() ?? throw new InvalidOperationException("No voices have been initiated."); }

        /// <summary>
        /// Get the header from a WAV stream
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        private static List<(string, byte[], object)> GetHeader(Stream stream)
        {
            byte[] header = new byte[46]; //create header byte array
            long pos = stream.Position;
            Console.WriteLine($"Stream position is {pos}");
            stream.Seek(0, SeekOrigin.Begin); 
            stream.Read(header, 0, 46); //
            List<byte> headerL = header.ToList();

            List <(string, byte[], object)> headerVals = new List<(string , byte[], object)>();

            //FILE TYPE CHUNK
            byte[] chunkId   = headerL.GetRange(0, 4).ToArray();
            headerVals.Add(("chunkID", chunkId, System.Text.Encoding.ASCII.GetString(chunkId)));
            byte[] chunkSize = headerL.GetRange(4, 4).ToArray();
            headerVals.Add(("chunkSize", chunkSize, BitConverter.ToInt32(chunkSize, 0)));
            byte[] format    = headerL.GetRange(8, 4).ToArray();
            headerVals.Add(("format", format, System.Text.Encoding.ASCII.GetString(format)));

            //FORMAT CHUNK
            byte[] formatChunkID   = headerL.GetRange(12, 4).ToArray();
            headerVals.Add(("formatChunkID", formatChunkID, System.Text.Encoding.ASCII.GetString(formatChunkID)));
            byte[] formatChunkSize = headerL.GetRange(16, 4).ToArray();
            headerVals.Add(("formatChunkSize", formatChunkSize, BitConverter.ToInt32(formatChunkSize, 0)));
            byte[] formatTag       = headerL.GetRange(20, 2).ToArray();
            headerVals.Add(("formatTag", formatTag, BitConverter.ToInt16(formatTag, 0)));
            byte[] numChannels     = headerL.GetRange(22, 2).ToArray();
            headerVals.Add(("numChannels", numChannels, BitConverter.ToInt16(numChannels, 0)));
            byte[] samplesPerSec   = headerL.GetRange(24, 4).ToArray();
            headerVals.Add(("samplesPerSec", samplesPerSec, BitConverter.ToInt32(samplesPerSec, 0)));
            byte[] avgBytesPerSec  = headerL.GetRange(28, 4).ToArray();
            headerVals.Add(("avgBytesPerSec", avgBytesPerSec, BitConverter.ToInt32(avgBytesPerSec, 0)));
            byte[] blockAlign      = headerL.GetRange(32, 2).ToArray();
            headerVals.Add(("blockAlign", blockAlign, BitConverter.ToInt16(blockAlign, 0)));
            byte[] bitsPerSample   = headerL.GetRange(34, 2).ToArray();
            headerVals.Add(("bitsPerSample", bitsPerSample, BitConverter.ToInt16(bitsPerSample, 0)));

            //DATA CHUNK
            byte[] dataChunkId   = header.ToList().GetRange(38, 4).ToArray();
            headerVals.Add(("dataChunkID", dataChunkId, System.Text.Encoding.ASCII.GetString(dataChunkId)));
            byte[] dataChunkSize = header.ToList().GetRange(42, 4).ToArray();
            headerVals.Add(("dataChunkSize", dataChunkSize, BitConverter.ToInt32(dataChunkSize, 0)));

            //Info logging
            Console.WriteLine("Infos : ");
            Console.WriteLine($" - Chunk ID          : {headerVals.FirstOrDefault(h => h.Item1 == "chunkID").Item3}");
            Console.WriteLine($" - ChunkSize         : {headerVals.FirstOrDefault(h => h.Item1 == "chunkSize").Item3}");
            Console.WriteLine($" - Format            : {headerVals.FirstOrDefault(h => h.Item1 == "format").Item3}");
            Console.WriteLine("");
            Console.WriteLine($" - Format Chunk ID   : {headerVals.FirstOrDefault(h => h.Item1 == "formatChunkID").Item3}");
            Console.WriteLine($" - Format Chunk Size : {headerVals.FirstOrDefault(h => h.Item1 == "formatChunkSize").Item3}");
            Console.WriteLine($" - Format tag        : {headerVals.FirstOrDefault(h => h.Item1 == "formatTag").Item3}");
            Console.WriteLine($" - Num Channels      : {headerVals.FirstOrDefault(h => h.Item1 == "numChannels").Item3}");
            Console.WriteLine($" - Samples Per Sec   : {headerVals.FirstOrDefault(h => h.Item1 == "samplesPerSec").Item3}");
            Console.WriteLine($" - Avg Bytes Per Sec : {headerVals.FirstOrDefault(h => h.Item1 == "avgBytesPerSec").Item3}");
            Console.WriteLine($" - Block Align       : {headerVals.FirstOrDefault(h => h.Item1 == "blockAlign").Item3}");
            Console.WriteLine($" - Bits per Sample   : {headerVals.FirstOrDefault(h => h.Item1 == "bitsPerSample").Item3}");
            Console.WriteLine("");
            Console.WriteLine($" - Data Chunk ID     : {headerVals.FirstOrDefault(h => h.Item1 == "dataChunkID").Item3}");
            Console.WriteLine($" - Data Chunk Size   : {headerVals.FirstOrDefault(h => h.Item1 == "dataChunkSize").Item3}");
            Console.WriteLine("");

            headerVals.Add(("fullHeader", header, null));
            stream.Seek(pos, SeekOrigin.Begin);
            return headerVals;
        }

        /// <summary>
        /// Parse header to byte array and stream
        /// </summary>
        /// <param name="headerValues"></param>
        /// <returns></returns>
        private static (byte[], Stream) SetHeader(List<(string, byte[], object)> headerValues)
        {
            byte[] header = new byte[46];
            foreach (var array in headerValues)
            {
                switch (array.Item1)
                {
                    case "chunkID":
                        array.Item2.CopyTo(header, 0);
                        break;
                    case "chunkSize":
                        array.Item2.CopyTo(header, 4);
                        break;
                    case "format":
                        array.Item2.CopyTo(header, 8);
                        break;
                    case "formatChunkID":
                        array.Item2.CopyTo(header, 12);
                        break;
                    case "formatChunkSize":
                        array.Item2.CopyTo(header, 16);
                        break;
                    case "formatTag":
                        array.Item2.CopyTo(header, 20);
                        break;
                    case "numChannels":
                        array.Item2.CopyTo(header, 22);
                        break;
                    case "samplesPerSec":
                        array.Item2.CopyTo(header, 24);
                        break;
                    case "avgBytesPerSec":
                        array.Item2.CopyTo(header, 28);
                        break;
                    case "blockAlign":
                        array.Item2.CopyTo(header, 32);
                        break;
                    case "bitsPerSample":
                        array.Item2.CopyTo(header, 34);
                        break;
                    case "dataChunkID":
                        array.Item2.CopyTo(header, 38);
                        break;
                    case "dataChunkSize":
                        array.Item2.CopyTo(header, 42);
                        break;
                }
            }
            Stream s = new MemoryStream(header);
            return (header, s);
        }

        /// <summary>
        /// Get a value in the header's list
        /// </summary>
        /// <param name="header"></param>
        /// <param name="query"></param>
        /// <param name="getParsed"></param>
        /// <returns></returns>
        private static object GetValueFromHeader(List<(string, byte[], object)> header, string query, bool getParsed = true)
        {
            if (getParsed) return header.FirstOrDefault(h => h.Item1 == query).Item3;
            return header.FirstOrDefault(h => h.Item1 == query).Item2;
        }

        /// <summary>
        /// Get a value's index in the Header list
        /// </summary>
        /// <param name="header"></param>
        /// <param name="query"></param>
        /// <returns></returns>
        private static int GetIndexFromHeader(List<(string, byte[], object)> header, string query)
        { return header.IndexOf(header.FirstOrDefault(h => h.Item1 == query)); }

    }
}
