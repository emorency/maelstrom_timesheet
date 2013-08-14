import gspread
import datetime
import argparse
import sys


def add_arguments(parser):
    """
    Add variable command specific options
    """
    # parser.add_argument('project',
    #                     help='Project name: [bioshare, clsa, cpac, bbmri, ialsa, cihr, interconnect, mrc, p3g, maelstrom]')
    parser.add_argument('--coordination', '-c', required=False, help='Coordination and Meetings')
    parser.add_argument('--dev_other', '-dev', required=False, help='Software Development Other')
    parser.add_argument('--datashield', '-d', required=False, help='DataShield Software Development')
    parser.add_argument('--opal', '-op', required=False, help='Opal Software Development')
    parser.add_argument('--onyx', '-on', required=False, help='Onyx Software Development')
    parser.add_argument('--mica', '-m', required=False, help='Mica Software Development')
    parser.add_argument('--day', type=int, required=False,
                        help='Day of the month (enter a number less than 0 to enter time for a previous day)')
    parser.add_argument('--month', required=False, help='Month of the year')
    # parser.add_argument('--yes', '-y', type=Boolean, required=False, help='Automatically answer yes')
    # parser.add_argument('--interactive', '-i', required=False, help='Interactive mode')


def add_sub_arguments(name):
    """
    Add sub commands
    """


def do_bioshare_command(args):
    args.project = "BioShare"
    args.col = 3
    do_command(args)


def do_clsa_command(args):
    args.project = "CLSA"
    args.col = 5
    do_command(args)


def do_maelstrom_command(args):
    args.project = "Maelstrom"
    args.col = 13
    do_command(args)


def do_p3g_command(args):
    args.project = "P3G"
    args.col = 12
    do_command(args)


def do_cpac_command(args):
    args.project = "CPAC"
    args.col = 6
    do_command(args)


def do_bbmri_command(args):
    args.project = "BBMRI"
    args.col = 7
    do_command(args)


def do_ialsa_command(args):
    args.project = "IALSA"
    args.col = 8
    do_command(args)


def do_cihr_command(args):
    args.project = "CIHR"
    args.col = 9
    do_command(args)


def do_inter_command(args):
    args.project = "InterConnect"
    args.col = 10
    do_command(args)


def do_mrc_command(args):
    args.project = "MRC"
    args.col = 11
    do_command(args)


def do_command(args):
    """
    Execute variable command
    """
    # Build and send request
    try:
        now = datetime.datetime.now()
        day = now.day
        month = now.month
        year = now.year

        if args.day and args.day > 0:
            day = args.day
        elif args.day and args.day < 0:
            # day = day + args.day
            now = now + datetime.timedelta(days=args.day)
            day = now.day
            month = now.month
            year = now.year

        if args.month:
            month = args.month

        currdate = '%s/%s/%s' % (month, day, year)

        print("Opening spreadsheet and finding info for date %s "  % currdate)

        # Login with your Google account
        gc = gspread.login('EMAIL', 'PASSWORD')
        sh = gc.open_by_key('0Ai3lwpfy7yHwdFBUNExYN3FoRGxjT1YwUXFrQ0JTX1E')
        wks = sh.get_worksheet(SHEET_NUMBER)

        # if not args.y:
        cell = wks.find(currdate)
        hours = wks.cell(cell.row + 10, cell.col + 1).value
        if not hours:
            hours = 0;

        print("Update project " + args.project + " (currently has %s hours) ? (y/n)" % hours)
        confirmed = sys.stdin.readline().rstrip().strip()
        if confirmed == "n":
            print("Aborted")
            sys.exit(2)

        row = cell.row

        # TASK
        if args.coordination:
            print("Coordination = " + args.coordination)
            wks.update_cell(row + 3, args.col, args.coordination)

        if args.dev_other:
            print("Dev Other = " + args.dev_other)
            wks.update_cell(row + 4, args.col, args.dev_other)

        if args.datashield:
            print("Datashield = " + args.datashield)
            wks.update_cell(row + 5, args.col, args.datashield)

        if args.opal:
            print("Opal = " + args.opal)
            wks.update_cell(row + 6, args.col, args.opal)

        if args.onyx:
            print("Onyx = " + args.onyx)
            wks.update_cell(row + 7, args.col, args.onyx)

        if args.mica:
            print("Mica = " + args.mica)
            wks.update_cell(row + 8, args.col, args.mica)

    except Exception, e:
        print
        e


def add_subcommand(name, help, add_args_func, default_func):
    """
    Make a sub-parser, add default arguments to it, add sub-command arguments and set the sub-command callback function.
    """
    subparser = subparsers.add_parser(name, help=help)
    add_arguments(subparser)
    # add_args_func(subparser)
    subparser.set_defaults(func=default_func)


# Parse arguments
parser = argparse.ArgumentParser(description='Timesheet command line.')
subparsers = parser.add_subparsers(title='sub-commands',
                                   help='Available sub-commands. Use --help option on the sub-command '
                                        'for more details.')

# Add subcommands
add_subcommand('maelstrom', 'Maelstrom', add_sub_arguments, do_maelstrom_command)
add_subcommand('bioshare', 'BioShare', add_sub_arguments, do_bioshare_command)
add_subcommand('p3g', 'P3G', add_sub_arguments, do_p3g_command)
add_subcommand('clsa', 'CLSA', add_sub_arguments, do_clsa_command)
add_subcommand('cpac', 'CPAC', add_sub_arguments, do_cpac_command)
add_subcommand('bbmri', 'BBMRI', add_sub_arguments, do_bbmri_command)
add_subcommand('ialsa', 'IALSA', add_sub_arguments, do_ialsa_command)
add_subcommand('cihr', 'CIHR', add_sub_arguments, do_cihr_command)
add_subcommand('interconnect', 'InterConnect', add_sub_arguments, do_inter_command)
add_subcommand('mrc', 'MRC', add_sub_arguments, do_mrc_command)

# Execute selected command
args = parser.parse_args()
args.func(args)
