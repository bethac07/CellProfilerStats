CellProfiler Pipeline: http://www.cellprofiler.org
Version:1
SVNRevision:11710

LoadImages:[module_num:1|svn_version:\'11587\'|variable_revision_number:11|show_window:False|notes:\x5B\x5D]
    File type to be loaded:individual images
    File selection method:Text-Exact match
    Number of images in each group?:3
    Type the text that the excluded images have in common:Do not use
    Analyze all subfolders within the selected folder?:None
    Input image file location:Default Input Folder\x7CNone
    Check image sets for missing or duplicate files?:Yes
    Group images by metadata?:No
    Exclude certain files?:No
    Specify metadata fields to group by:
    Select subfolders to analyze:
    Image count:1
    Text that these images have in common (case-sensitive):tif
    Position of this image in each group:1
    Extract metadata from where?:None
    Regular expression that finds metadata in the file name:^(?P<Plate>.*)_(?P<Well>\x5BA-P\x5D\x5B0-9\x5D{2})_s(?P<Site>\x5B0-9\x5D)
    Type the regular expression that finds metadata in the subfolder path:.*\x5B\\\\/\x5D(?P<Date>.*)\x5B\\\\/\x5D(?P<Run>.*)$
    Channel count:1
    Group the movie frames?:No
    Grouping method:Interleaved
    Number of channels per group:3
    Load the input as images or objects?:Images
    Name this loaded image:DNA
    Name this loaded object:Nuclei
    Retain outlines of loaded objects?:No
    Name the outline image:LoadedImageOutlines
    Channel number:1
    Rescale intensities?:Yes

ColorToGray:[module_num:2|svn_version:\'10318\'|variable_revision_number:2|show_window:False|notes:\x5B\x5D]
    Select the input image:DNA
    Conversion method:Split
    Image type\x3A:Channels
    Name the output image:OrigGray
    Relative weight of the red channel:1
    Relative weight of the green channel:1
    Relative weight of the blue channel:1
    Convert red to gray?:Yes
    Name the output image:OrigRed
    Convert green to gray?:Yes
    Name the output image:OrigGreen
    Convert blue to gray?:Yes
    Name the output image:OrigBlue
    Channel count:2
    Channel number\x3A:Red\x3A 1
    Relative weight of the channel:1
    Image name\x3A:Channel1
    Channel number\x3A:Green\x3A 2
    Relative weight of the channel:1
    Image name\x3A:Channel2

GrayToColor:[module_num:3|svn_version:\'10341\'|variable_revision_number:2|show_window:False|notes:\x5B\x5D]
    Select a color scheme:RGB
    Select the input image to be colored red:Channel1
    Select the input image to be colored green:Leave this black
    Select the input image to be colored blue:Channel2
    Name the output image:ColorImagea
    Relative weight for the red image:1
    Relative weight for the green image:1
    Relative weight for the blue image:1
    Select the input image to be colored cyan:Leave this black
    Select the input image to be colored magenta:Leave this black
    Select the input image to be colored yellow:Leave this black
    Select the input image that determines brightness:Leave this black
    Relative weight for the cyan image:1
    Relative weight for the magenta image:1
    Relative weight for the yellow image:1
    Relative weight for the brightness image:1
    Select the input image to add to the stacked image:None

CorrectIlluminationCalculate:[module_num:4|svn_version:\'10458\'|variable_revision_number:2|show_window:False|notes:\x5B\x5D]
    Select the input image:Channel1
    Name the output image:IllumRed
    Select how the illumination function is calculated:Background
    Dilate objects in the final averaged image?:No
    Dilation radius:1
    Block size:10
    Rescale the illumination function?:No
    Calculate function for each image individually, or based on all images?:Each
    Smoothing method:Splines
    Method to calculate smoothing filter size:Object size
    Approximate object size:300
    Smoothing filter size:10
    Retain the averaged image for use later in the pipeline (for example, in SaveImages)?:No
    Name the averaged image:IllumBlueAvg
    Retain the dilated image for use later in the pipeline (for example, in SaveImages)?:No
    Name the dilated image:IllumBlueDilated
    Automatically calculate spline parameters?:No
    Background mode:auto
    Number of spline points:10
    Background threshold:2
    Image resampling factor:2
    Maximum number of iterations:40
    Residual value for convergence:0.001

CorrectIlluminationCalculate:[module_num:5|svn_version:\'10458\'|variable_revision_number:2|show_window:False|notes:\x5B\x5D]
    Select the input image:Channel2
    Name the output image:IllumBlue
    Select how the illumination function is calculated:Background
    Dilate objects in the final averaged image?:No
    Dilation radius:1
    Block size:10
    Rescale the illumination function?:No
    Calculate function for each image individually, or based on all images?:Each
    Smoothing method:Fit Polynomial
    Method to calculate smoothing filter size:Automatic
    Approximate object size:10
    Smoothing filter size:10
    Retain the averaged image for use later in the pipeline (for example, in SaveImages)?:No
    Name the averaged image:IllumBlueAvg
    Retain the dilated image for use later in the pipeline (for example, in SaveImages)?:No
    Name the dilated image:IllumBlueDilated
    Automatically calculate spline parameters?:Yes
    Background mode:auto
    Number of spline points:5
    Background threshold:2
    Image resampling factor:2
    Maximum number of iterations:40
    Residual value for convergence:0.001

CorrectIlluminationApply:[module_num:6|svn_version:\'10300\'|variable_revision_number:3|show_window:True|notes:\x5B\x5D]
    Select the input image:Channel2
    Name the output image:CorrBlue
    Select the illumination function:IllumBlue
    Select how the illumination function is applied:Subtract
    Select the input image:Channel1
    Name the output image:CorrRed
    Select the illumination function:IllumRed
    Select how the illumination function is applied:Subtract

Smooth:[module_num:7|svn_version:\'10465\'|variable_revision_number:1|show_window:False|notes:\x5B\x5D]
    Select the input image:CorrRed
    Name the output image:FilteredImage
    Select smoothing method:Gaussian Filter
    Calculate artifact diameter automatically?:No
    Typical artifact diameter, in  pixels:25
    Edge intensity difference:0.1

GrayToColor:[module_num:8|svn_version:\'10341\'|variable_revision_number:2|show_window:False|notes:\x5B\x5D]
    Select a color scheme:RGB
    Select the input image to be colored red:CorrRed
    Select the input image to be colored green:Leave this black
    Select the input image to be colored blue:CorrBlue
    Name the output image:ColorImage
    Relative weight for the red image:1
    Relative weight for the green image:1
    Relative weight for the blue image:1
    Select the input image to be colored cyan:Leave this black
    Select the input image to be colored magenta:Leave this black
    Select the input image to be colored yellow:Leave this black
    Select the input image that determines brightness:Leave this black
    Relative weight for the cyan image:1
    Relative weight for the magenta image:1
    Relative weight for the yellow image:1
    Relative weight for the brightness image:1
    Select the input image to add to the stacked image:None

IdentifyPrimaryObjects:[module_num:9|svn_version:\'10826\'|variable_revision_number:8|show_window:False|notes:\x5B\x5D]
    Select the input image:Channel1
    Name the primary objects to be identified:UnCorrNuclei
    Typical diameter of objects, in pixel units (Min,Max):150,400
    Discard objects outside the diameter range?:Yes
    Try to merge too small objects with nearby larger objects?:No
    Discard objects touching the border of the image?:No
    Select the thresholding method:RidlerCalvard Global
    Threshold correction factor:1.7
    Lower and upper bounds on threshold:0.000000,1.000000
    Approximate fraction of image covered by objects?:0.01
    Method to distinguish clumped objects:Intensity
    Method to draw dividing lines between clumped objects:Intensity
    Size of smoothing filter:10
    Suppress local maxima that are closer than this minimum allowed distance:7
    Speed up by using lower-resolution image to find local maxima?:Yes
    Name the outline image:UnCorrOutlines
    Fill holes in identified objects?:Yes
    Automatically calculate size of smoothing filter?:Yes
    Automatically calculate minimum allowed distance between local maxima?:Yes
    Manual threshold:0.0
    Select binary image:None
    Retain outlines of the identified objects?:Yes
    Automatically calculate the threshold using the Otsu method?:Yes
    Enter Laplacian of Gaussian threshold:0.5
    Two-class or three-class thresholding?:Two classes
    Minimize the weighted variance or the entropy?:Weighted variance
    Assign pixels in the middle intensity class to the foreground or the background?:Foreground
    Automatically calculate the size of objects for the Laplacian of Gaussian filter?:Yes
    Enter LoG filter diameter:5
    Handling of objects if excessive number of objects identified:Continue
    Maximum number of objects:500
    Select the measurement to threshold with:None

IdentifyPrimaryObjects:[module_num:10|svn_version:\'10826\'|variable_revision_number:8|show_window:True|notes:\x5B\x5D]
    Select the input image:FilteredImage
    Name the primary objects to be identified:SmoothNuclei
    Typical diameter of objects, in pixel units (Min,Max):150,400
    Discard objects outside the diameter range?:Yes
    Try to merge too small objects with nearby larger objects?:No
    Discard objects touching the border of the image?:No
    Select the thresholding method:RidlerCalvard Global
    Threshold correction factor:1
    Lower and upper bounds on threshold:0.000000,1.000000
    Approximate fraction of image covered by objects?:0.01
    Method to distinguish clumped objects:Intensity
    Method to draw dividing lines between clumped objects:Intensity
    Size of smoothing filter:10
    Suppress local maxima that are closer than this minimum allowed distance:7
    Speed up by using lower-resolution image to find local maxima?:Yes
    Name the outline image:SmoothOutlines
    Fill holes in identified objects?:Yes
    Automatically calculate size of smoothing filter?:Yes
    Automatically calculate minimum allowed distance between local maxima?:Yes
    Manual threshold:0.0
    Select binary image:None
    Retain outlines of the identified objects?:Yes
    Automatically calculate the threshold using the Otsu method?:Yes
    Enter Laplacian of Gaussian threshold:0.5
    Two-class or three-class thresholding?:Two classes
    Minimize the weighted variance or the entropy?:Weighted variance
    Assign pixels in the middle intensity class to the foreground or the background?:Foreground
    Automatically calculate the size of objects for the Laplacian of Gaussian filter?:Yes
    Enter LoG filter diameter:5
    Handling of objects if excessive number of objects identified:Continue
    Maximum number of objects:500
    Select the measurement to threshold with:None

IdentifyPrimaryObjects:[module_num:11|svn_version:\'10826\'|variable_revision_number:8|show_window:False|notes:\x5B\x5D]
    Select the input image:CorrRed
    Name the primary objects to be identified:CorrNuclei
    Typical diameter of objects, in pixel units (Min,Max):150,400
    Discard objects outside the diameter range?:Yes
    Try to merge too small objects with nearby larger objects?:No
    Discard objects touching the border of the image?:No
    Select the thresholding method:RidlerCalvard Global
    Threshold correction factor:1.7
    Lower and upper bounds on threshold:0.000000,1.000000
    Approximate fraction of image covered by objects?:0.01
    Method to distinguish clumped objects:Intensity
    Method to draw dividing lines between clumped objects:Intensity
    Size of smoothing filter:50
    Suppress local maxima that are closer than this minimum allowed distance:7
    Speed up by using lower-resolution image to find local maxima?:Yes
    Name the outline image:CorrOutlines
    Fill holes in identified objects?:Yes
    Automatically calculate size of smoothing filter?:Yes
    Automatically calculate minimum allowed distance between local maxima?:Yes
    Manual threshold:0.0
    Select binary image:None
    Retain outlines of the identified objects?:Yes
    Automatically calculate the threshold using the Otsu method?:Yes
    Enter Laplacian of Gaussian threshold:0.5
    Two-class or three-class thresholding?:Two classes
    Minimize the weighted variance or the entropy?:Weighted variance
    Assign pixels in the middle intensity class to the foreground or the background?:Foreground
    Automatically calculate the size of objects for the Laplacian of Gaussian filter?:Yes
    Enter LoG filter diameter:5
    Handling of objects if excessive number of objects identified:Continue
    Maximum number of objects:500
    Select the measurement to threshold with:None

OverlayOutlines:[module_num:12|svn_version:\'10672\'|variable_revision_number:2|show_window:True|notes:\x5B\x5D]
    Display outlines on a blank image?:No
    Select image on which to display outlines:ColorImage
    Name the output image:OrigOverlay
    Select outline display mode:Color
    Select method to determine brightness of outlines:Max of image
    Width of outlines:1
    Select outlines to display:UnCorrOutlines
    Select outline color:Red
    Select outlines to display:CorrOutlines
    Select outline color:White
    Select outlines to display:SmoothOutlines
    Select outline color:Blue

SaveImages:[module_num:13|svn_version:\'10822\'|variable_revision_number:7|show_window:False|notes:\x5B\x5D]
    Select the type of image to save:Image
    Select the image to save:OrigOverlay
    Select the objects to save:None
    Select the module display window to save:None
    Select method for constructing file names:From image filename
    Select image name for file prefix:DNA
    Enter single file name:OrigBlue
    Do you want to add a suffix to the image file name?:No
    Text to append to the image name:
    Select file format to use:bmp
    Output file location:Default Output Folder\x7CNone
    Image bit depth:8
    Overwrite existing files without warning?:No
    Select how often to save:Every cycle
    Rescale the images? :Yes
    Save as grayscale or color image?:Grayscale
    Select colormap:gray
    Store file and path information to the saved image?:No
    Create subfolders in the output folder?:No
