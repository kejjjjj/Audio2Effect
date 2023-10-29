using System;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Drawing;
using System.Runtime;
using System.Xml;
using ScriptPortal.Vegas;

public class EntryPoint
{
    float var_LUFSTriggerValue = 0.0f;
    float var_effectDurationValue = 0.0f;
    bool var_overlappingEffects = false;
    bool var_showAverageAudios = false;
    bool create_gui()
    {
        // Create the main form
        var form = new Form
        {
            Text = "Trigger Settings",
            Width = 300,
            Height = 200
        };

        // Create the "dB trigger" label and input
        var label1 = new Label
        {
            Text = "LUFS trigger:",
            Left = 10,
            Top = 20,
            Width = 100
        };
        var dBTriggerInput = new NumericUpDown
        {
            Left = 120,
            Top = 15,
            Width = 150,
            //DecimalPlaces = 2,
            Maximum = 0,
            Minimum = -60,
            Value = -20
        };

        // Create the "effect duration" label and input
        var label2 = new Label
        {
            Text = "Effect Duration:",
            Left = 10,
            Top = 50,
            Width = 100
        };
        var effectDurationInput = new NumericUpDown
        {
            Left = 120,
            Top = 45,
            Width = 150,
            DecimalPlaces = 2,
            Minimum = 0,
            Value = 0.2M
        };
        var overlappingEffectsCheckBox = new CheckBox
        {
            Text = "Overlapping Effects",
            Left = 10,
            Top = 80,
            Width = 200
        };
        var showAveragesCheckBox = new CheckBox
        {
            Text = "Show average audio",
            Left = 10,
            Top = 100,
            Width = 200
        };
        // Create the "OK" button
        var okButton = new Button
        {
            Text = "OK",
            Left = 120,
            Top = 130,
            Width = 75
        };

        // Event handler for the OK button click
        okButton.Click += (sender, e) =>
        {
            var_LUFSTriggerValue = (float)dBTriggerInput.Value;
            var_effectDurationValue = (float)effectDurationInput.Value;
            var_overlappingEffects = overlappingEffectsCheckBox.Checked;
            var_showAverageAudios = showAveragesCheckBox.Checked;

            if (var_LUFSTriggerValue >= 0.0f)
            {
                throw new Exception("LUFS Trigger must be less than 0");
            }
            if (var_effectDurationValue <= 0.0f)
            {
                throw new Exception("effect duration must be greater than 0");
            }

            form.DialogResult = DialogResult.OK;
            form.Close();
        };

        // Add the controls to the form
        form.Controls.Add(label1);
        form.Controls.Add(dBTriggerInput);
        form.Controls.Add(label2);
        form.Controls.Add(effectDurationInput);
        form.Controls.Add(overlappingEffectsCheckBox);
        form.Controls.Add(showAveragesCheckBox);
        form.Controls.Add(okButton);

        dBTriggerInput.Select();
        form.AcceptButton = okButton;
        return form.ShowDialog() == DialogResult.OK;
    }

    public class Loudness_Data
    {

        public Loudness_Data(Timecode _tc, float _volume)
        {
            tc = _tc;
            volume = _volume;
        }

        public Timecode tc;
        public float volume;
    }

    List<TrackEvent> get_active_tracks(Project project)
    {
        List<TrackEvent> selected_tracks = new List<TrackEvent>();
        foreach (Track track in project.Tracks)
        {
            foreach (TrackEvent trackEvent in track.Events)
            {
                if (trackEvent.Selected)
                    selected_tracks.Add(trackEvent);
            }
        }
        return selected_tracks;
    }

    List<TrackEvent> trim_closest_to_beginning(TrackEvent a, TrackEvent b)
    {
        var a_ms = a.Start.ToMilliseconds();
        var b_ms = b.Start.ToMilliseconds();

        if (a.Start == b.Start)
            return new List<TrackEvent>() { a, b };

        TrackEvent closest_to_0 = a_ms < b_ms ? a : b;
        closest_to_0 = closest_to_0.Split(a_ms < b_ms ? (b.Start - a.Start) : (a.Start - b.Start));

        return new List<TrackEvent>() { closest_to_0, a_ms < b_ms ? b : a };

    }
    List<TrackEvent> trim_closest_to_end(TrackEvent a, TrackEvent b)
    {
        var a_ms = a.End.ToMilliseconds();
        var b_ms = b.End.ToMilliseconds();

        if (a_ms == b_ms)
            return new List<TrackEvent>() { a, b };

        TrackEvent closest_to_end = a_ms > b_ms ? a : b;

        var length = closest_to_end.End - closest_to_end.Start;
        closest_to_end.Split(a_ms < b_ms ? length - (b.End - a.End) : length - (a.End - b.End));
        return new List<TrackEvent>() { closest_to_end, a_ms > b_ms ? b : a };
    }

    Tuple<Renderer, RenderTemplate> get_renderer_and_template(Vegas vegas, string classID, string renderTemplateGuid)
    {
        Renderer _renderer = null;
        RenderTemplate template = null;

        foreach (Renderer renderer in vegas.Renderers)
        {
            if (classID == renderer.ClassID.ToString())
            {
                _renderer = renderer;

                foreach (RenderTemplate renderTemplate in renderer.Templates)
                {
                    if (renderTemplate.IsValid())
                    {
                        if (renderTemplateGuid == renderTemplate.TemplateGuid.ToString())
                        {
                            template = renderTemplate;
                        }
                    }
                }
            }
        }

        if (_renderer == null)
        {
            throw new Exception("couldn't find appropriate renderer");
        }
        else if (template == null)
        {
            throw new Exception("couldn't find appropriate template");
        }
        return new Tuple<Renderer, RenderTemplate>(_renderer, template);
    }

    void render_temp_file(Vegas vegas, TrackEvent track, string temp_file, RenderTemplate template)
    {
        RenderArgs args = new RenderArgs();
        args.OutputFile = temp_file;
        args.RenderTemplate = template;
        args.Start = track.Start;
        args.Length = track.Length;
        args.IncludeMarkers = false;
        args.StretchToFill = false;
        args.GenerateLoudnessLog = true;

        RenderStatus status = vegas.Render(args);

        if (status == RenderStatus.Failed || status == RenderStatus.Quit || status == RenderStatus.Canceled || status == RenderStatus.Unknown)
        {
            throw new Exception("the rendering failed");
        }
    }

    List<List<Loudness_Data>> parse_loudness(string file)
    {
        IEnumerable<String> lines;
        try
        {
            lines = File.ReadLines(file);
        }
        catch (Exception ex)
        {
            throw ex;
        }
        bool header_parsed = false;
        var loudness_segments = new List<List<Loudness_Data>>();
        var loudness_list = new List<Loudness_Data>();

        foreach (string line in lines)
        {
            if (line.StartsWith("------------"))
            {
                break;
            }
            else if (line.StartsWith("              Pos."))
            {
                if (!header_parsed)
                {
                    header_parsed = true;
                }
                continue;
            }
            if (header_parsed == false)
                continue;

            if (line.Length > 5)
            {
                string[] pieces = line.Split('\t');
                Timecode tc = Timecode.FromString(pieces[1]);
                float volume;
                if (pieces[2].Contains("Inf"))
                {
                    volume = -100;
                }
                else
                {
                    volume = (float)Convert.ToDouble(pieces[2], CultureInfo.InvariantCulture);
                }
                if (volume > -40.0f)
                {
                    loudness_list.Add(new Loudness_Data(tc, volume));

                    if (loudness_list.Count > 1000)
                    {
                        loudness_segments.Add(new List<Loudness_Data>(loudness_list));
                        loudness_list.Clear();
                    }
                }
            }

        }

        if (loudness_list.Count > 0)
        {
            loudness_segments.Add(new List<Loudness_Data>(loudness_list));
        }
        return loudness_segments;

    }

    float get_average_volume(List<Loudness_Data> loudness_data)
    {
        float sum = 0.0F;

        foreach (Loudness_Data data in loudness_data)
            sum += data.volume;

        return sum / loudness_data.Count;

    }

    Timecode zoom_in_effect(Vegas vegas, VideoEvent videoEvent, Timecode start)
    {
        float duration = var_effectDurationValue;

        VideoMotionKeyframe key1 = new VideoMotionKeyframe(start);
        VideoMotionKeyframe key2 = new VideoMotionKeyframe(start + Timecode.FromSeconds(duration / 2));
        VideoMotionKeyframe key3 = new VideoMotionKeyframe(start + Timecode.FromSeconds(duration));

        videoEvent.VideoMotion.Keyframes.Add(key1);
        videoEvent.VideoMotion.Keyframes.Add(key2);
        videoEvent.VideoMotion.Keyframes.Add(key3);

        int numKeys = videoEvent.VideoMotion.Keyframes.Count;

        VideoMotionKeyframe vkey1 = videoEvent.VideoMotion.Keyframes[numKeys - 3];
        VideoMotionKeyframe vkey2 = videoEvent.VideoMotion.Keyframes[numKeys - 2];
        VideoMotionKeyframe vkey3 = videoEvent.VideoMotion.Keyframes[numKeys - 1];

        int videoWidth = vegas.Project.Video.Width;
        int videoHeight = vegas.Project.Video.Width;


        vkey1.ScaleBy(new VideoMotionVertex(1.0f, 1.0f));
        vkey2.ScaleBy(new VideoMotionVertex(0.8f, 0.8f));
        vkey3.ScaleBy(new VideoMotionVertex(1.0f, 1.0f));

        return key3.Position;
    }

    Timecode apply_effect_if_loud(Vegas vegas, VideoEvent videoEvent, Loudness_Data data, float average_volume)
    {
        if ((data.volume < var_LUFSTriggerValue))
            return null;

        return zoom_in_effect(vegas, videoEvent, data.tc);
    }

    public void FromVegas(Vegas vegas)
    {
        try
        {

            if (create_gui() == false)
                return;

            List<TrackEvent> tracks = get_active_tracks(vegas.Project);
            if (tracks.Count != 2)
            {
                throw new Exception("must have 2 tracks selected");
            }

            if (tracks[0].MediaType == MediaType.Unknown || tracks[1].MediaType == MediaType.Unknown
                || tracks[0].MediaType == tracks[1].MediaType)
            {
                throw new Exception("must have an audio track and video track selected");
            }

            Tuple<Renderer, RenderTemplate> r = get_renderer_and_template(vegas, "adfa6a4b-a99b-42e3-ae1f-081123ada04b", "8ab64a16-81f5-46e6-8155-1611d592088c");

            tracks = trim_closest_to_beginning(tracks[0], tracks[1]);
            tracks = trim_closest_to_end(tracks[0], tracks[1]);

            TrackEvent audioTrack = tracks[0].IsAudio() ? tracks[0] : tracks[1];
            TrackEvent videoTrack = tracks[0].IsVideo() ? tracks[0] : tracks[1];
            VideoEvent videoEvent = (VideoEvent)videoTrack;

            render_temp_file(vegas, audioTrack, vegas.TemporaryFilesPath + "\\temp.mp3", r.Item2);

            List<List<Loudness_Data>> data_all = parse_loudness(vegas.TemporaryFilesPath + "\\temp_loud.txt");

            string average_audios = "a list of average audios between segments: \n";

            foreach (List<Loudness_Data> data in data_all)
            {


                float average_audio = get_average_volume(data);

                // MessageBox.Show("average volume: " + Convert.ToString(average_audio));
                Timecode next_allowed_effect = data[0].tc;

                foreach (Loudness_Data d in data)
                {
                    if (d.tc < next_allowed_effect && var_overlappingEffects == false)
                        continue;

                    Timecode next = apply_effect_if_loud(vegas, videoEvent, d, average_audio);

                    if (next != null)
                    {
                        next_allowed_effect = next;
                    }
                }

                string first = data[0].tc.ToString();
                string last = data[data.Count - 1].tc.ToString();

                
                   average_audios += "| " + first + ", " + last +  " | -> " + Convert.ToString(average_audio) + "\n";
            }
            if (var_showAverageAudios)
                MessageBox.Show(average_audios);
        }

        catch (Exception exception)
        {
            MessageBox.Show("Error: " + exception);
        }
        

    }
}