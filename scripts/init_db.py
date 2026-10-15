from factory_core import initialize_database, load_settings, record_run


def main() -> None:
    settings = load_settings()
    initialize_database(settings.database_path)
    run_id = record_run(settings.database_path, source="init_db")
    print(f"Base SQLite inicializada: {settings.database_path}")
    print(f"Ejecución registrada: {run_id}")


if __name__ == "__main__":
    main()

