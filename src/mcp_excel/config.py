"""Configuration settings for MCP Excel Server."""

from pydantic import Field
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    """Application settings loaded from environment variables."""

    # Transport settings
    transport: str = Field(default="stdio", description="Transport type: 'stdio' or 'http'")

    # Security settings
    mask_errors: bool = Field(default=True, description="Mask internal errors in production")

    # Cache settings
    cache_max_size: int = Field(default=5, description="Maximum number of workbooks in cache")
    cache_max_memory_mb: int = Field(default=1024, description="Maximum cache memory in MB")

    # Pagination settings
    paging_cells_limit: int = Field(default=4000, description="Maximum cells per page")

    # Response limits
    max_response_rows: int = Field(default=500, description="Maximum rows in response")
    max_response_columns: int = Field(default=50, description="Maximum columns in response")

    # Logging
    log_level: str = Field(default="INFO", description="Logging level")

    model_config = {
        "env_prefix": "MCP_EXCEL_",
        "env_file": ".env",
        "env_file_encoding": "utf-8",
    }


# Global settings instance
settings = Settings()
