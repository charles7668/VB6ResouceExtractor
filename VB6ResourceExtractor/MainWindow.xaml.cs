using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using Image = System.Drawing.Image;
using MessageBox = System.Windows.Forms.MessageBox;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;

namespace VB6ResourceExtractor;

/// <summary>
///     Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow
{
    private const int ImageKey = 0x746C; // work for little-endian
    private const int ListKey = 0x010003; // work for little-endian

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
    }

    public ObservableCollection<ResourceItem> ResourceItems { get; } = new();

    private void BtnOpenBinary_OnClick(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "frx|*.frx",
            Title = "select vb6 binary file"
        };
        if (openFileDialog.ShowDialog() == false ||
            string.IsNullOrEmpty(openFileDialog.FileName))
            return;
        ResourceItems.Clear();
        var fileName = openFileDialog.FileName;
        var fileContent = File.ReadAllBytes(fileName);
        var offset = 0;
        while (offset < fileContent.Length)
        {
            // check if data is image
            var key = BitConverter.ToInt32(fileContent, offset + 4);
            if (key == ImageKey)
            {
                var size = BitConverter.ToInt32(fileContent, offset + 8);
                var imageContent = new byte[size];
                Array.Copy(fileContent, offset + 12, imageContent, 0, size);
                using (var ms = new MemoryStream(imageContent))
                {
                    var image = Image.FromStream(ms);
                    ResourceItems.Add(new ResourceItem(image.RawFormat.ToString(), ResourceType.Image, imageContent));
                    image.Dispose();
                }

                offset += size + 12;
            }
            else
            {
                // if resource is for the list control (e.g. combobox, list) , then ItemData and ListData have same header
                // header : 2byte is length , 4 byte is 0x00010003
                // if length is 0 , then header is 00 00 , without 0x00010003
                var listLength = BitConverter.ToInt16(fileContent, offset);
                if (listLength == 0)
                {
                    offset += 2;
                }
                else
                {
                    var listKey = BitConverter.ToInt32(fileContent, offset + 2);
                    if (listKey != ListKey)
                    {
                        MessageBox.Show("convert failed, unknown type");
                        return;
                    }

                    var patternList = new List<(byte[] content, int shift)>
                    {
                        (new byte[] { 0x03, 0x00, 0x01, 0x00 }, -2),
                        (new byte[] { 0x6c, 0x74, 0x00, 0x00 }, -4),
                        (new byte[] { 0x00, 0x00 }, 0)
                    };
                    var endIndex = offset + 6;
                    var isMatch = false;
                    foreach (var pattern in patternList)
                    {
                        for (var i = offset + 6; i < fileContent.Length; i++)
                        {
                            isMatch = fileContent.Skip(i).Take(pattern.content.Length)
                                .SequenceEqual(pattern.content);
                            if (!isMatch)
                                continue;
                            endIndex = i + pattern.shift;
                            break;
                        }

                        if (isMatch)
                            break;
                    }

                    if (!isMatch)
                    {
                        endIndex = fileContent.Length;
                    }
                    else
                    {
                        var length = endIndex - offset - 6;
                        var listContent = new byte[length];
                        Array.Copy(fileContent, offset + 6, listContent, 0, length);
                        ResourceItems.Add(new ResourceItem("ListItem", ResourceType.ListItem, listContent));
                    }

                    offset = endIndex;
                }
            }
        }
    }

    private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
    {
        var index = ListResources.SelectedIndex;
        if (index < 0)
            return;
        var resourceItem = ResourceItems[index];
        if (resourceItem.ResourceType != ResourceType.Image)
            return;
        var saveFileDialog = new SaveFileDialog
        {
            Title = "save image",
            FileName = "image"
        };
        if (saveFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            return;
        var fileName = saveFileDialog.FileName;
        var memoryStream = new MemoryStream((byte[])resourceItem.Content);
        var image = Image.FromStream(memoryStream);
        image.Save(Path.Combine($"{fileName}.{image.RawFormat}"));
    }

    private void Selector_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        var listObject = (ListView)sender;
        var index = listObject.SelectedIndex;
        if (index < 0)
            return;
        var resourceItem = ResourceItems[index];

        if (resourceItem.ResourceType == ResourceType.Image)
        {
            var memoryStream = new MemoryStream((byte[])resourceItem.Content);
            // Create a new BitmapImage and set its source to the MemoryStream
            memoryStream.Position = 0;
            var bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.StreamSource = memoryStream;
            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
            bitmapImage.EndInit();
            bitmapImage.Freeze();

            // Set the ImageSource of your image control to the BitmapImage
            ImgImageObject.Source = bitmapImage;
            SetResourceVisibility(true);
        }
        else
        {
            TxtStringObject.Text = Encoding.UTF8.GetString((byte[])resourceItem.Content).Replace('\0', '\n');
            SetResourceVisibility(false);
        }
    }

    private void SetResourceVisibility(bool isImage)
    {
        ImgImageObject.Visibility = isImage ? Visibility.Visible : Visibility.Hidden;
        ScrollTxtStringObject.Visibility = isImage ? Visibility.Hidden : Visibility.Visible;
    }
}