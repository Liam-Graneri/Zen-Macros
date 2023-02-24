import os
import csv
import re
import shutil
import random
import json
import subprocess
import datetime
import sys
from System.IO import Path, Directory

# Get current time
now = datetime.datetime.now().strftime('%Y-%m-%d_%H%M')

#################### Functions ####################
###################################################
# Creates a metadata_dict object with empty values blah blah:


def metadata_dict(setup):

    return {
        'analysed': [],
        'half_analysed': [],
        'prefix': '',
        'regions': [],
        'blind_names': {},
        'analysis_began': now,
        'setup': setup
    }

# Updates metadata object


def update_metadata_json(metadata_json, analysed_files, hidden=True):
    try:
        os.remove(metadata_json)
    except:
        None
    with open(metadata_json, 'w') as md_file:
        json.dump(analysed_files, md_file)
    if hidden:
        subprocess.check_call(['attrib', '+H', metadata_json])

# Defines a Function for Cancel Button


def cancel_macro(end_message='', current_image={}):
    if current_image != {}:
        try:
            update_metadata_json(analysed_files)
        except:
            None
    if end_message == '':
        sys.exit('Macro Aborted with Cancel.')
    else:
        sys.exit(end_message)

# Creates an option window in Zen


def create_zen_window(header, content_dict, cancel_message='', *args, **kwargs):
    # Create Base Window
    window = ZenWindow()
    window.Initialize(header)
    # Add details from content_dict
    elements_pointers = []
    keys = list(content_dict.keys())
    keys.sort()
    for key in keys:
        content = [x for x in content_dict[key]]
        if 'label' in key:
            window.AddLabel(content_dict[key])
        else:
            if 'folderBrowser' in key:
                window.AddFolderBrowser(*content)
            elif 'dropDown' in key:
                window.AddDropDown(*content)
            elif 'checkBox' in key:
                window.AddCheckbox(*content)
            elif 'textBox' in key:
                window.AddTextBox(*content)
            elif 'integerRange' in key:
                window.AddIntegerRange(*content)
            else:
                continue
            elements_pointers.append(content[0])
    init_window = window.Show()
    if init_window.HasCanceled == True:
        cancel_macro(end_message=cancel_message, *args, **kwargs)

    elements_inputs = {}
    for variable in elements_pointers:
        elements_inputs[variable] = str(init_window.GetValue(variable))

    return elements_inputs

# Create a prefix using setup information


def create_prefix(setup):
    prefix = re.sub(
        ' ', '-', re.sub(
            '_{2,}', '_', '_'.join([
                setup['expNo'], setup['sample'], setup['stain']
            ])
        )
    )
    prefix = prefix + '_' if prefix != '_' else ''
    return prefix

# Blinding myself with some randomised mad libs


def randomise_name(analysed_files):
    adjectives = ['Amazing', 'Backwards', 'Bearded', 'Blatant', 'Boring', 'Bumbling', 'Cheesy', 'Chunky', 'Clumsy', 'Comedic', 'Confused', 'Creamy', 'Crisp', 'Crispified', 'Danish', 'Delicious', 'Desperate', 'Domesticated', 'Dry', 'Extra', 'Fishy', 'Flooded', 'Foamy', 'Friendly', 'Frigid', 'Gasping', 'Glittering', 'Hard', 'Helpless', 'Hungry', 'HUUUGE', 'Iconic', 'Improper', 'Inedible', 'Innocent', 'Insolent', 'Jovial', 'Lazy', 'Leaking', 'Lush', 'Magical', 'Milky', 'Misty',
                  'Moist', 'Motherly', 'Nasty', 'Nefarious', 'Obese', 'Offensive', 'Overpowered', 'Overt', 'Peculiar', 'Precious', 'Proper', 'Rolling', 'Rude', 'Salty', 'Shiny', 'Skilled', 'Smooth', 'Sniffing', 'Soft', 'Sour', 'Sticky', 'Subtle', 'Sulking', 'Talkative', 'Tropical', 'Unsubtle', 'Vague', 'Untouched', 'Wet', 'Witty', 'Vivid', 'Illegal', 'Loose', 'Upset', 'Unwashed', 'Desolate', 'Squeaky', 'Submissive', 'Gushing', 'Willing', 'Suave', 'Cartoony', 'Crunchy', 'Suspicious']
    nouns = ['Arms', 'Bagels', 'Bagpipes', 'Balls', 'Bamboos', 'Belltowers', 'Bits', 'Bones', 'Bosses', 'Bushes', 'Cakes', 'Carrots', 'Chipmunks', 'Chunks', 'Clementines', 'Dogs', 'Dustbunnies', 'Ears', 'Eyes', 'Feet', 'Fingers', 'Firemen', 'Fish', 'Fossils', 'Geeks', 'Brains', 'Hats', 'Heads', 'Herbs', 'Hipsters', 'Housewives', 'Icecubes', 'Minerals', 'Mushrooms', 'Noses', 'Nuts',
             'Packages', 'Pants', 'Organisms', 'Pearls', 'Pieces', 'Presidents', 'Pronks', 'Puppies', 'Rangers', 'Rodents', 'Sausages', 'Scouts', 'Shovels', 'Soup', 'Submarines', 'Teabags', 'Tentacles', 'Truckers', 'Turtles', 'Warriors', 'Wigs', 'Cucumbers', 'Aunts', 'Abs', 'Bananas', 'Gems', 'Muffins', 'Monks', 'Craters', 'Mouths', 'Eggplants', 'Peaches', 'Kilts', 'Teapots', 'Sailors']
    adjective = random.choice(adjectives)
    noun = random.choice(nouns)
    random_name = "The_{}_{}".format(adjective.title(), noun.title())
    while random_name in analysed_files['blind_names'].values():
        random_name = randomise_name(analysed_files)
    return random_name

###################################################
###################################################


# Warn users that this is a custom macro. There may be issues.
Zen.Application.Pause('This macro was written by Liam Graneri. If you use this and something goes wrong, I take no responsibility.\nIf something is broken with this macro and needs fixing, please come bearing gifts.')
home = os.path.expanduser('~')


# Setup Dialogue
setup = create_zen_window(
    header='Batch Analyse Interactively',
    content_dict={
        '01-label': ('---    Select Image(s) Folder    ---'),
        '02-folderBrowser': ('sourcefolder', 'Source folder with images'),
        '03-dropDown': ('extension', 'Input image type', ['*.czi', '*.zvi', '.lsm', '*.jpeg', '.jpg', '*.png', '*.tiff', '*.tif'], 0),
        '06-folderBrowser': ('destfolder', 'Select folder for ROI output'),
        '07-checkBox': ('multiple-regions', 'Analyse multiple regions per scene?\t', True),
        '08-checkBox': ('blind', 'Perform all analyses blind?\t\t', True),
        '09-label': ('---    Optional Output Name Params    ---'),
        '10-textBox': ('expNo', 'Experiment #', None),
        '11-textBox': ('sample', 'Sample type', None),
        '12-textBox': ('stain', 'Stain', None),
    }
)

# Check there are scan in 'input' folder
files = Directory.GetFiles(setup['sourcefolder'], setup['extension'])
if files.Length == 0:
    cancel_macro('There are no images of type: {} in the selected directory').format(
        setup['extension'])

# Check if returning to analysis
returning = False
metadata_json = Path.Combine(
    setup['destfolder'], 'zen_split_metadata.json')
if os.path.exists(metadata_json):
    analysed_files = {}
    with open(metadata_json) as md_file:
        md = json.load(md_file)
        analysed_files.update(md)
    return_string = ('--- Previous Params ---\n'
                     'Exp No:\t\t{}\n'
                     'Sample:\t\t{}\n'
                     'Stain:\t\t{}\n\n'
                     '--- Regions ---\n'
                     '{}\n\n'
                     '--- Analysis Began ---\n'
                     '{}').format(
        analysed_files['setup']['expNo'],
        analysed_files['setup']['sample'],
        analysed_files['setup']['stain'],
        '\n'.join(analysed_files['regions']),
        datetime.datetime.strptime(
            analysed_files['analysis_began'], '%Y-%m-%d_%H%M').strftime('%c')
    )
    returning = create_zen_window(
        header='Returning from previous analysis?',
        content_dict={'01-checkBox': ('returning', return_string, True)}
    )
    if returning['returning'] != 'True':
        analysed_files = metadata_dict(setup)
        update_metadata_json(metadata_json, analysed_files)
        returning = False
    else:
        analysed_files['setup'].update(
            {'sourcefolder': setup['sourcefolder'],
                'destfolder': setup['destfolder']}
        )
        setup = analysed_files['setup']
        returning = True
else:
    analysed_files = metadata_dict(setup)
    update_metadata_json(metadata_json, analysed_files)


# Get prefix from setup data
prefix = create_prefix(setup)
analysed_files['prefix'] = prefix if not returning else analysed_files.get(
    'prefix')
# Get remaining files to analyse
already_analysed = [Path.Combine(
    setup['sourcefolder'], x+setup['extension'][1:]) for x in analysed_files['analysed']]
files = [x for x in files if x not in already_analysed]
random.shuffle(files) if setup['blind'] == 'True' else files




# Allows User to Label Multi-Regions if Selected
if setup['multiple-regions'] == 'True' and not returning:
    # Asks user how many regions of interest they want to specify per image
    regionsSpecify = create_zen_window(
        header='Define the number of regions per image.',
        content_dict={
            '01-label': ('---    ROI Number    ---'),
            '02-integerRange': ('nroi', 'How may regions of interest in each scene?', 1, 1, 5)
        })
    nROI = int(regionsSpecify['nroi'])
    # Asks the user to label each region
    resultDict = {'01-label': ('---    Title Individual ROIs    ---')}
    for n in range(nROI):
        element = n + 2
        if element < 10:
            resultDict[str(element)+'-textBox'] = ('roiName' +
                                                   str(n+1), 'Region '+str(n+1)+' Name', None)
        else:
            resultDict[str(element)+'-textBox'] = ('roiName' +
                                                   str(n+1), 'Region '+str(n+1)+' Name', None)
    labelRegions = create_zen_window(
        header='Label ROIs',
        content_dict=resultDict)
    regions = []
    labels = labelRegions.keys()
    labels.sort()
    for label in labels:
        regions.append(labelRegions[label])
    analysed_files['regions'] = regions
elif setup['multiple-regions'] != 'True':
    regions = ['Not specified']
else:
    regions = analysed_files['regions']
update_metadata_json(metadata_json, analysed_files)

# Loops over all images in the chosen directory
for file_index, file in enumerate(files):
    # Loads image into Zen
    file_path = Path.Combine(setup['sourcefolder'], file)
    image = Zen.Application.LoadImage(file_path, False)
    Zen.Application.Documents.Add(image)
    original_filename = Path.GetFileNameWithoutExtension(image.Name)
    random_name = analysed_files['blind_names'].get(
            original_filename, randomise_name(analysed_files))
    if setup['blind'] == 'True':
        image.Name = random_name
        analysed_files['blind_names'][original_filename] = random_name
    # Loop through regions
    for region in regions:
        Zen.Application.Pause('Add ROIs for {} region'.format(region))
        for i in range(0, image.Graphics.Count):
            graphic = image.Graphics[i]
            graphic.IsMeasurementVisible = False
            # Get coordinates of bounding box
            XS = graphic.Bounds.Left - 10
            YS = graphic.Bounds.Top - 10
            XE = graphic.Bounds.Right + 10
            YE = graphic.Bounds.Bottom + 10
            #Create bounding box image of ROI
            str_boundbox = 'X({}-{})|Y({}-{})'.format(int(XS),int(XE),int(YS),int(YE))
            boundbox = Zen.Processing.Utilities.CreateSubset(image, str_boundbox)
            Zen.Application.Documents.Add(boundbox)
            #Save bounding box subset image
            boundbox_filename = '{}_{}_ROI-{}.czi'.format(original_filename, region, i+1)
            boundbox_path = Path.Combine(setup['destfolder'], boundbox_filename)
            boundbox.Save(boundbox_path)
            boundbox.Close()
        image.Close()
        image = Zen.Application.LoadImage(file_path, False)
        Zen.Application.Documents.Add(image)
        image.Name = random_name if setup['blind'] == 'True' else None
    analysed_files['analysed'].append(original_filename)
    update_metadata_json(metadata_json, analysed_files)
    image.Close()
    remaining_files = len(files) - file_index - 1
    if file_index < len(files) - 1:
        update_metadata_json(metadata_json, analysed_files)
        continue_analysing = create_zen_window(
            header='Continue Analysing Slides?',
            content_dict={
                '01-checkBox': ('cont', 'Remaining: {}'.format(remaining_files), True)}
        )
        if continue_analysing['cont'] != 'True':
            break
        else:
            continue
    else:
        try:
            shutil.move(metadata_json, final_results_dir)
        except:
            None