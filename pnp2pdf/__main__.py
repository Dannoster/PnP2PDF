import sys


def main(argv=None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    if argv:
        from pnp2pdf import cli
        return cli.main(argv)
    from pnp2pdf.app import main as app_main
    app_main()
    return 0


if __name__ == "__main__":
    sys.exit(main())
