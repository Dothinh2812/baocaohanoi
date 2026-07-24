"""Provider interface cho AI question generation.

Provider production là adapter cấu hình qua env; mặc định AI disabled.
Test dùng fake provider. Không gọi API AI thật trong test.
"""


class GenerationResult:
    def __init__(self, batch, *, provider, model=None, prompt_version=None,
                 usage=None, raw_response=None):
        self.batch = batch
        self.provider = provider
        self.model = model
        self.prompt_version = prompt_version
        self.usage = usage
        self.raw_response = raw_response


class BaseProvider:
    name = "base"

    def generate(self, *, source_document_version_ids, target_audience_codes,
                 requested_count, **kwargs):
        raise NotImplementedError
