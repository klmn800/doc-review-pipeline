"""
Document Review Pipeline

A configurable pipeline for batch-reviewing documents with an LLM. Scans a
folder for Word documents (and optionally email files containing .docx
attachments), extracts the text, sends each document to Claude with a
customizable evaluation prompt, and writes structured results to an Excel
tracker.

The evaluation criteria are defined entirely in the prompts/ folder, so the
same pipeline can review any class of document by editing the prompt files.
The default shipped prompts evaluate policy documents for the presence of
both policy language and procedural guidance, but any review criteria can
be substituted - see prompts/examples/ for alternatives.

License: MIT
"""

import os
import sys
import csv
import json
import re
import shutil
import time
import configparser
import email
from email import policy
from email.parser import BytesParser
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Optional

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill
except ImportError:
    print("ERROR: openpyxl not found. Install with: pip install openpyxl")
    sys.exit(1)

try:
    from docx import Document
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph
except ImportError:
    print("ERROR: python-docx not found. Install with: pip install python-docx")
    sys.exit(1)

try:
    import anthropic
except ImportError:
    print("ERROR: anthropic not found. Install with: pip install anthropic")
    sys.exit(1)

try:
    import extract_msg
except ImportError:
    print("ERROR: extract-msg not found. Install with: pip install extract-msg")
    sys.exit(1)

try:
    from dateutil import parser as date_parser
except ImportError:
    print("ERROR: python-dateutil not found. Install with: pip install python-dateutil")
    sys.exit(1)


def is_file_locked(filepath: str) -> bool:
    """
    Check if a file is locked by another process (e.g., open in Excel).

    Args:
        filepath: Path to the file to check

    Returns:
        True if file is locked, False if available
    """
    if not os.path.exists(filepath):
        return False

    try:
        # Try to open with exclusive write access
        with open(filepath, 'a'):
            return False
    except (IOError, PermissionError):
        return True


class DocReviewPipeline:
    """Main orchestration class for the document review pipeline."""

    # Pricing per million tokens (input/output) for different Claude models
    MODEL_PRICING = {
        'claude-sonnet-4-5-20250929': {'input': 3.0, 'output': 15.0},
        'claude-haiku-4-5-20251001': {'input': 0.80, 'output': 4.0},
        'claude-opus-4-5-20251101': {'input': 15.0, 'output': 75.0},
    }

    def __init__(self, config_path: str = "config.ini"):
        """
        Initialize the pipeline with configuration.

        Args:
            config_path: Path to the configuration file
        """
        self.config_path = config_path
        self.config = self.load_config()
        self.log_messages = []
        self.current_log_file = None
        self.log_timestamp = None

    def load_config(self) -> configparser.ConfigParser:
        """
        Load configuration from config.ini file.

        Returns:
            ConfigParser object with loaded settings

        Raises:
            SystemExit: If config file is missing or invalid
        """
        config = configparser.ConfigParser()

        if not os.path.exists(self.config_path):
            print(f"ERROR: Config file not found at {self.config_path}")
            print("Please create config.ini with your API key and settings.")
            sys.exit(1)

        config.read(self.config_path)

        # Validate required sections
        required_sections = ['API', 'Paths', 'Settings']
        for section in required_sections:
            if section not in config:
                print(f"ERROR: Missing [{section}] section in config.ini")
                sys.exit(1)

        # Check for API key
        api_key = config['API'].get('anthropic_api_key', '').strip()
        if not api_key or api_key == 'your_key_here':
            print("ERROR: Please set your anthropic_api_key in config.ini")
            print("Get your API key at: https://console.anthropic.com/")
            sys.exit(1)

        return config

    def setup_log_file(self) -> str:
        """
        Create a timestamped log file for this run.

        Initially creates as "check_" - will be renamed to "review_" if work is found.

        Returns:
            Path to the created log file
        """
        log_folder = self.config['Paths'].get('log_folder', './logs')
        os.makedirs(log_folder, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_folder, f"check_{timestamp}.txt")

        self.current_log_file = log_file
        self.log_timestamp = timestamp
        return log_file

    def rename_log_to_review(self):
        """
        Rename the current log file from check_ to review_ prefix.
        Called when documents are found to process.
        """
        if self.current_log_file and os.path.exists(self.current_log_file):
            log_folder = self.config['Paths'].get('log_folder', './logs')
            new_log_file = os.path.join(log_folder, f"review_{self.log_timestamp}.txt")

            try:
                os.rename(self.current_log_file, new_log_file)
                self.current_log_file = new_log_file
            except Exception as e:
                self.log(f"Warning: Could not rename log file: {str(e)}", to_console=False)

    def log(self, message: str, to_console: bool = True):
        """
        Log a message to both console and log file.

        Args:
            message: The message to log
            to_console: Whether to also print to console (default: True)
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}"

        self.log_messages.append(log_entry)

        if to_console:
            print(message)

        # Write to log file if it exists
        if self.current_log_file:
            with open(self.current_log_file, 'a', encoding='utf-8') as f:
                f.write(log_entry + "\n")

    def scan_folder(self) -> List[str]:
        """
        Scan the input folder for .docx files.

        Returns:
            List of paths to .docx files found
        """
        input_folder = self.config['Paths'].get('input_folder', './inbox')

        if not os.path.exists(input_folder):
            self.log(f"ERROR: Input folder not found: {input_folder}")
            return []

        docx_files = []
        for file in os.listdir(input_folder):
            if file.endswith('.docx') and not file.startswith('~$'):
                full_path = os.path.join(input_folder, file)
                docx_files.append(full_path)

        return docx_files

    def scan_email_files(self) -> List[str]:
        """
        Scan the input folder for .msg and .eml email files.

        Returns:
            List of paths to email files found (.msg and .eml)
        """
        input_folder = self.config['Paths'].get('input_folder', './inbox')

        if not os.path.exists(input_folder):
            return []

        email_files = []
        for file in os.listdir(input_folder):
            if file.endswith('.msg') or file.endswith('.eml'):
                full_path = os.path.join(input_folder, file)
                email_files.append(full_path)

        return email_files

    def extract_text(self, docx_path: str) -> Tuple[Optional[str], Optional[str]]:
        """
        Extract text from a Word document, preserving numbered/bulleted lists.

        Args:
            docx_path: Path to the .docx file

        Returns:
            Tuple of (extracted_text, error_message)
            If successful, returns (text, None)
            If failed, returns (None, error_message)
        """
        try:
            doc = Document(docx_path)

            # Accept tracked changes if configured
            accept_changes = self.config['Settings'].getboolean('accept_tracked_changes', True)
            if accept_changes:
                # Note: python-docx doesn't have direct API for this
                # Tracked changes are automatically accepted when reading
                pass

            preserve_lists = self.config['Settings'].getboolean('preserve_lists', True)

            extracted_parts = []

            for element in doc.element.body:
                if isinstance(element, CT_P):
                    paragraph = Paragraph(element, doc)
                    text = paragraph.text.strip()

                    if not text:
                        continue

                    # Check if paragraph has list formatting
                    if preserve_lists and paragraph.style.name.startswith(('List', 'Numbered')):
                        # Try to preserve list structure
                        num_format = self._get_list_format(paragraph)
                        if num_format:
                            extracted_parts.append(f"{num_format} {text}")
                        else:
                            extracted_parts.append(f"• {text}")
                    else:
                        extracted_parts.append(text)

                elif isinstance(element, CT_Tbl):
                    # Extract table content
                    table = Table(element, doc)
                    table_text = self._extract_table_text(table)
                    if table_text:
                        extracted_parts.append(table_text)

            full_text = "\n\n".join(extracted_parts)

            if not full_text.strip():
                return None, "Document appears to be empty"

            return full_text, None

        except Exception as e:
            return None, f"Failed to extract text: {str(e)}"

    def _get_list_format(self, paragraph) -> Optional[str]:
        """
        Try to extract the list number/bullet from a paragraph.

        Args:
            paragraph: python-docx Paragraph object

        Returns:
            List format string (e.g., "1.", "a)", "•") or None
        """
        # This is a simplified approach - proper list numbering is complex
        # We'll just check if it starts with a number or letter
        text = paragraph.text.strip()
        if text and len(text) > 2:
            if text[0].isdigit() and text[1] in '.):':
                return text[:2]
            elif text[0].isalpha() and text[1] in '.):':
                return text[:2]
        return None

    def _extract_table_text(self, table: Table) -> str:
        """
        Extract text from a Word table.

        Args:
            table: python-docx Table object

        Returns:
            Formatted table text
        """
        rows = []
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if any(cells):  # Only include non-empty rows
                rows.append(" | ".join(cells))

        return "\n".join(rows) if rows else ""

    def load_prompts(self) -> Tuple[str, str]:
        """
        Load system prompt and instruction files.

        Returns:
            Tuple of (system_prompt, instructions_text)
        """
        prompts_folder = self.config['Paths'].get('prompts_folder', './prompts')

        # Load main system prompt
        system_prompt_path = os.path.join(prompts_folder, 'system_prompt.txt')
        if not os.path.exists(system_prompt_path):
            self.log(f"WARNING: System prompt not found at {system_prompt_path}")
            system_prompt = "You are a document reviewer."
        else:
            with open(system_prompt_path, 'r', encoding='utf-8') as f:
                system_prompt = f.read()

        # Load all instruction files
        instructions_folder = os.path.join(prompts_folder, 'instructions')
        instructions_parts = []

        if os.path.exists(instructions_folder):
            for filename in sorted(os.listdir(instructions_folder)):
                if filename.endswith('.txt'):
                    filepath = os.path.join(instructions_folder, filename)
                    with open(filepath, 'r', encoding='utf-8') as f:
                        instructions_parts.append(f.read())

        instructions_text = "\n\n".join(instructions_parts) if instructions_parts else ""

        return system_prompt, instructions_text

    def calculate_cost(self, model: str, input_tokens: int, output_tokens: int) -> float:
        """
        Calculate the cost based on model and token usage.

        Args:
            model: The Claude model being used
            input_tokens: Number of input tokens
            output_tokens: Number of output tokens

        Returns:
            Total cost in USD
        """
        # Get pricing for the model, default to Sonnet pricing if unknown
        pricing = self.MODEL_PRICING.get(model, {'input': 3.0, 'output': 15.0})

        input_cost = (input_tokens / 1_000_000) * pricing['input']
        output_cost = (output_tokens / 1_000_000) * pricing['output']

        return input_cost + output_cost

    def extract_summary(self, body_text: str) -> str:
        """
        Extract the project summary from email body.

        Looks for text between summary headers and "Comment Due Date:".
        Handles variations: "Project Summary:", "Summary of Draft:", "Summary:", etc.

        Args:
            body_text: The email body text

        Returns:
            Extracted summary text, or empty string if not found
        """
        try:
            # Try multiple summary header variations
            patterns = [
                r'Summary of Draft:\s*(.*?)\s*Comment Due Date:',      # "Summary of Draft:"
                r'Project Summary:\s*(.*?)\s*Comment Due Date:',       # "Project Summary:"
                r'Summary:\s*(.*?)\s*Comment Due Date:',               # "Summary:"
                r'Draft Summary:\s*(.*?)\s*Comment Due Date:',         # "Draft Summary:"
            ]

            for pattern in patterns:
                match = re.search(pattern, body_text, re.IGNORECASE | re.DOTALL)
                if match:
                    summary = match.group(1).strip()
                    return summary

            return ""
        except Exception as e:
            self.log(f"  Warning: Could not extract summary: {str(e)}", to_console=False)
            return ""

    def extract_due_date(self, body_text: str) -> str:
        """
        Extract and parse the due date from email body.

        Uses fuzzy date parsing to handle various formats.

        Args:
            body_text: The email body text

        Returns:
            Due date in YYYY-MM-DD format, or empty string if not found
        """
        try:
            # Look for text after "Comment Due Date:" - capture first line only
            # This handles various line ending styles and avoids multi-line issues
            pattern = r'Comment Due Date:\s*([^\n\r]+)'
            match = re.search(pattern, body_text, re.IGNORECASE)

            if match:
                due_date_text = match.group(1).strip()

                # Use dateutil's fuzzy parser to extract date
                parsed_date = date_parser.parse(due_date_text, fuzzy=True)
                return parsed_date.strftime("%Y-%m-%d")

            return ""
        except Exception as e:
            self.log(f"  Warning: Could not parse due date: {str(e)}", to_console=False)
            return ""

    def parse_email_file(self, msg_path: str) -> Tuple[Optional[Dict], Optional[str]]:
        """
        Parse a .msg email file to extract metadata and attachments.

        Args:
            msg_path: Path to the .msg file

        Returns:
            Tuple of (email_metadata_dict, error_message)
            If successful, returns (metadata, None)
            If failed, returns (None, error_message)
        """
        try:
            msg = extract_msg.Message(msg_path)

            # Extract sender name (prefer sender name over email address)
            sender = msg.sender or "Unknown Sender"

            # Extract subject and optionally remove "FOR YOUR REVIEW:" prefix
            subject = msg.subject or "No Subject"
            # Remove "FOR YOUR REVIEW:" prefix (case insensitive, with optional punctuation)
            subject = re.sub(r'^for\s+your\s+review:?\s*', '', subject, flags=re.IGNORECASE).strip()

            # Extract body text
            body = msg.body or ""

            # Extract summary and due date from body
            summary = self.extract_summary(body)
            due_date = self.extract_due_date(body)

            # Extract .docx attachments
            attachments = []
            temp_dir = os.path.join(self.config['Paths'].get('input_folder', './inbox'), '_temp_attachments')
            os.makedirs(temp_dir, exist_ok=True)

            # Debug: log all attachments found
            total_attachments = len(msg.attachments)
            self.log(f"  Debug: Found {total_attachments} total attachment(s) in .msg file", to_console=False)

            for i, attachment in enumerate(msg.attachments, 1):
                filename = attachment.longFilename or attachment.shortFilename
                self.log(f"  Debug: Attachment {i}: {filename}", to_console=False)

                if filename and filename.lower().endswith('.docx'):
                    # Save attachment to temp directory
                    attachment_path = os.path.join(temp_dir, filename)
                    with open(attachment_path, 'wb') as f:
                        f.write(attachment.data)
                    attachments.append(attachment_path)
                    self.log(f"  Debug: Extracted .docx: {filename}", to_console=False)

            msg.close()

            metadata = {
                'sender': sender,
                'subject': subject,
                'summary': summary,
                'due_date': due_date,
                'attachments': attachments,
                'msg_file': msg_path
            }

            return metadata, None

        except Exception as e:
            return None, f"Failed to parse email: {str(e)}"

    def parse_eml_file(self, eml_path: str) -> Tuple[Optional[Dict], Optional[str]]:
        """
        Parse a .eml email file to extract metadata and attachments.

        Args:
            eml_path: Path to the .eml file

        Returns:
            Tuple of (email_metadata_dict, error_message)
            If successful, returns (metadata, None)
            If failed, returns (None, error_message)
        """
        try:
            with open(eml_path, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)

            # Extract sender
            from_header = msg.get('From', 'Unknown Sender')
            # Try to extract just the name part if it's in "Name <email>" format
            sender_match = re.match(r'^"?(.+?)"?\s*<.+>$', from_header)
            sender = sender_match.group(1) if sender_match else from_header

            # Extract subject and remove "FOR YOUR REVIEW:" prefix
            subject = msg.get('Subject', 'No Subject')
            subject = re.sub(r'^for\s+your\s+review:?\s*', '', subject, flags=re.IGNORECASE).strip()

            # Extract body text (prefer plain text, fallback to HTML)
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == 'text/plain':
                        body = part.get_content()
                        break
                    elif content_type == 'text/html' and not body:
                        # Fallback to HTML if no plain text
                        body = part.get_content()
            else:
                body = msg.get_content()

            # Extract summary and due date from body
            summary = self.extract_summary(body)
            due_date = self.extract_due_date(body)

            # Extract .docx attachments
            attachments = []
            temp_dir = os.path.join(self.config['Paths'].get('input_folder', './inbox'), '_temp_attachments')
            os.makedirs(temp_dir, exist_ok=True)

            attachment_count = 0
            for part in msg.walk():
                # Check if this part is an attachment
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    attachment_count += 1
                    self.log(f"  Debug: Attachment {attachment_count}: {filename}", to_console=False)

                    if filename and filename.lower().endswith('.docx'):
                        # Save attachment to temp directory
                        attachment_path = os.path.join(temp_dir, filename)
                        with open(attachment_path, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        attachments.append(attachment_path)
                        self.log(f"  Debug: Extracted .docx: {filename}", to_console=False)

            self.log(f"  Debug: Found {attachment_count} total attachment(s) in .eml file", to_console=False)

            metadata = {
                'sender': sender,
                'subject': subject,
                'summary': summary,
                'due_date': due_date,
                'attachments': attachments,
                'msg_file': eml_path
            }

            return metadata, None

        except Exception as e:
            return None, f"Failed to parse .eml file: {str(e)}"

    def analyze_document(self, text: str, filename: str) -> Dict:
        """
        Send document text to Claude API for analysis.

        Args:
            text: Extracted document text
            filename: Name of the file being analyzed

        Returns:
            Dictionary with analysis results including:
            - review_status: "Passed", "Needs Review", or "Error"
            - review_notes: Explanation from Claude
            - tokens_used: Number of tokens consumed
            - cost: Estimated cost in USD
            - error: Error message if applicable
        """
        api_key = self.config['API'].get('anthropic_api_key')
        model = self.config['API'].get('model', 'claude-sonnet-4-5-20250929')
        max_tokens = self.config['API'].getint('max_tokens', 2000)

        system_prompt, instructions = self.load_prompts()

        # Construct the full prompt
        user_prompt = f"""Document: {filename}

{instructions}

---

Document Text:

{text}

---

Please analyze this document and provide your review."""

        try:
            client = anthropic.Anthropic(api_key=api_key)

            response = client.messages.create(
                model=model,
                max_tokens=max_tokens,
                system=system_prompt,
                messages=[
                    {"role": "user", "content": user_prompt}
                ]
            )

            # Extract response
            response_text = response.content[0].text

            # Calculate tokens and cost
            input_tokens = response.usage.input_tokens
            output_tokens = response.usage.output_tokens
            total_tokens = input_tokens + output_tokens

            # Calculate cost based on the model being used
            total_cost = self.calculate_cost(model, input_tokens, output_tokens)

            # Try to parse JSON response
            try:
                # Look for JSON in the response
                if '{' in response_text and '}' in response_text:
                    json_start = response_text.index('{')
                    json_end = response_text.rindex('}') + 1
                    json_str = response_text[json_start:json_end]
                    result = json.loads(json_str)

                    return {
                        'review_status': result.get('review_status', 'Needs Review'),
                        'review_notes': result.get('review_notes', response_text),
                        'tokens_used': total_tokens,
                        'cost': total_cost,
                        'error': None
                    }
                else:
                    # Fallback if no JSON found
                    return {
                        'review_status': 'Needs Review',
                        'review_notes': response_text,
                        'tokens_used': total_tokens,
                        'cost': total_cost,
                        'error': None
                    }
            except json.JSONDecodeError:
                # If JSON parsing fails, use the raw response
                return {
                    'review_status': 'Needs Review',
                    'review_notes': response_text,
                    'tokens_used': total_tokens,
                    'cost': total_cost,
                    'error': None
                }

        except anthropic.APIError as e:
            return {
                'review_status': 'Error',
                'review_notes': f"API Error: {str(e)}",
                'tokens_used': 0,
                'cost': 0.0,
                'error': str(e)
            }
        except Exception as e:
            return {
                'review_status': 'Error',
                'review_notes': f"Unexpected error: {str(e)}",
                'tokens_used': 0,
                'cost': 0.0,
                'error': str(e)
            }

    def update_tracker(self, results: List[Dict]) -> str:
        """
        Write or update the Excel tracker file with analysis results.
        Appends new entries to existing tracker file.

        Args:
            results: List of result dictionaries from document analyses

        Returns:
            Path to the tracker file
        """
        # Get tracker path
        tracker_file = self.config['Paths'].get('tracker_file', './tracker.xlsx')

        # Check if file is locked (e.g., open in Excel)
        if os.path.exists(tracker_file) and is_file_locked(tracker_file):
            self.log("⚠ WARNING: tracker.xlsx is currently open in Excel")
            self.log("Retrying in 30 seconds... (close the file to avoid creating a backup)")

            # Retry 3 times with 30-second delays
            for attempt in range(1, 4):
                time.sleep(30)
                if not is_file_locked(tracker_file):
                    self.log("✓ File is now available - continuing with normal save")
                    break
                else:
                    if attempt < 3:
                        self.log(f"File still locked - retry {attempt + 1}/3 in 30 seconds...")
            else:
                # Still locked after 3 attempts - save to backup
                backup_file = tracker_file.replace('.xlsx', f'_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
                self.log(f"✗ File still locked after 90 seconds")
                self.log(f"Saving to backup file: {os.path.basename(backup_file)}")
                tracker_file = backup_file

        # Define column headers
        headers = [
            'Document Name',
            'Date Processed',
            'Review Status',
            'Review Notes',
            'Tokens Used',
            'Cost',
            'File Path',
            'Due Date',
            'Assigned To',
            'Summary',
            'Sender',
            'Email Subject'
        ]

        # Check if tracker file exists
        file_exists = os.path.exists(tracker_file)

        if file_exists:
            # Load existing workbook
            wb = load_workbook(tracker_file)
            ws = wb.active
        else:
            # Create new workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "Document Reviews"

            # Write header row with formatting
            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_num, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                cell.font = Font(bold=True, color="FFFFFF")

        # Append results
        for result in results:
            # Extract email metadata if available
            email_meta = result.get('email_metadata', None)

            if email_meta:
                due_date = email_meta.get('due_date', '')
                summary = email_meta.get('summary', '')
                sender = email_meta.get('sender', '')
                email_subject = email_meta.get('subject', '')
            else:
                due_date = ''
                summary = ''
                sender = ''
                email_subject = ''

            # Prepare row data
            row_data = [
                result.get('document_name', ''),
                result.get('date_processed', ''),
                result.get('review_status', ''),
                result.get('review_notes', ''),
                result.get('tokens_used', 0),
                f"${result.get('cost', 0.0):.4f}",
                result.get('file_path', ''),
                due_date,
                '',  # Assigned To (manual fill)
                summary,
                sender,
                email_subject
            ]

            ws.append(row_data)

        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            ws.column_dimensions[column_letter].width = adjusted_width

        # Save workbook
        wb.save(tracker_file)

        return tracker_file

    def move_to_processed(self, file_path: str) -> bool:
        """
        Move a processed file to the processed folder.

        Args:
            file_path: Path to the file to move

        Returns:
            True if successful, False otherwise
        """
        try:
            processed_folder = self.config['Paths'].get('processed_folder', './processed')
            os.makedirs(processed_folder, exist_ok=True)

            filename = os.path.basename(file_path)
            destination = os.path.join(processed_folder, filename)

            # Handle duplicate filenames
            if os.path.exists(destination):
                base, ext = os.path.splitext(filename)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"{base}_{timestamp}{ext}"
                destination = os.path.join(processed_folder, filename)

            shutil.move(file_path, destination)
            self.log(f"  ✓ Moved to processed folder: {filename}", to_console=False)
            return True

        except Exception as e:
            self.log(f"  ✗ Failed to move file: {str(e)}")
            return False

    def get_model_display_name(self, model: str) -> str:
        """
        Get a friendly display name for the model.

        Args:
            model: Model identifier

        Returns:
            Human-readable model name
        """
        model_names = {
            'claude-sonnet-4-5-20250929': 'Claude Sonnet 4.5',
            'claude-haiku-4-5-20251001': 'Claude Haiku 4.5',
            'claude-opus-4-5-20251101': 'Claude Opus 4.5',
        }
        return model_names.get(model, model)

    def run(self):
        """Main execution method - orchestrates the entire workflow."""

        # Setup logging
        log_file = self.setup_log_file()
        self.log("=" * 60)
        self.log("Document Review Pipeline")
        self.log("=" * 60)
        self.log(f"Log file: {log_file}")
        self.log("")

        # Phase 1: Process email files and extract attachments
        self.log("Scanning for email files...")
        msg_files = self.scan_email_files()

        email_metadata_map = {}  # Maps document path -> email metadata
        msg_files_to_move = []  # Track .msg files to move later

        if msg_files:
            self.log(f"Found {len(msg_files)} email file(s).\n")

            for msg_path in msg_files:
                msg_filename = os.path.basename(msg_path)
                self.log(f"Processing email: {msg_filename}")

                # Determine parser based on file extension
                if msg_path.lower().endswith('.eml'):
                    metadata, error = self.parse_eml_file(msg_path)
                else:
                    metadata, error = self.parse_email_file(msg_path)

                if error:
                    self.log(f"  ✗ Error: {error}")
                    continue

                # Log extracted metadata
                self.log(f"  From: {metadata['sender']}")
                self.log(f"  Subject: {metadata['subject'][:60]}...")
                self.log(f"  Due Date: {metadata['due_date'] or 'Not found'}")
                self.log(f"  Attachments: {len(metadata['attachments'])} document(s)")

                # Map each attachment to this email's metadata
                for attachment_path in metadata['attachments']:
                    email_metadata_map[attachment_path] = metadata

                msg_files_to_move.append(msg_path)
                self.log("")
        else:
            self.log("No email files found.")
            self.log("")

        # Phase 2: Scan for standalone documents
        self.log("Scanning for documents...")
        standalone_docx = self.scan_folder()

        # Combine attachments from emails + standalone documents
        all_documents = list(email_metadata_map.keys()) + standalone_docx

        if not all_documents:
            self.log("No documents found in inbox folder.")
            self.log("Add .docx files or .msg emails to the inbox folder and run again.")
            return

        # Documents found - rename log file from check_ to review_
        self.rename_log_to_review()

        self.log(f"Found {len(all_documents)} document(s) to process.\n")

        # Phase 3: Process each document
        results = []
        total_tokens = 0
        total_cost = 0.0

        for i, file_path in enumerate(all_documents, 1):
            filename = os.path.basename(file_path)
            self.log(f"Processing [{i}/{len(all_documents)}]: {filename}")

            # Extract text
            text, error = self.extract_text(file_path)

            if error:
                self.log(f"  ✗ Error: {error}")
                result = {
                    'document_name': filename,
                    'date_processed': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'review_status': 'Error',
                    'review_notes': error,
                    'tokens_used': 0,
                    'cost': 0.0,
                    'file_path': file_path,
                    'email_metadata': email_metadata_map.get(file_path, None)
                }
                results.append(result)
                continue

            # Analyze with Claude
            self.log(f"  Analyzing with Claude...")
            analysis = self.analyze_document(text, filename)

            # Prepare result with email metadata if available
            result = {
                'document_name': filename,
                'date_processed': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'review_status': analysis['review_status'],
                'review_notes': analysis['review_notes'],
                'tokens_used': analysis['tokens_used'],
                'cost': analysis['cost'],
                'file_path': file_path,
                'email_metadata': email_metadata_map.get(file_path, None)
            }

            results.append(result)
            total_tokens += analysis['tokens_used']
            total_cost += analysis['cost']

            # Log result
            status_icon = "✓" if analysis['review_status'] == 'Passed' else "⚠"
            self.log(f"  {status_icon} Status: {analysis['review_status']}")
            self.log(f"  Notes: {analysis['review_notes']}")
            self.log(f"  Tokens: {analysis['tokens_used']} | Cost: ${analysis['cost']:.4f}")

            # Move standalone documents to processed folder (attachments moved with cleanup)
            if analysis['review_status'] != 'Error' and file_path not in email_metadata_map:
                self.move_to_processed(file_path)

            self.log("")

        # Phase 4: Cleanup - move email files and temp attachments
        temp_dir = os.path.join(self.config['Paths'].get('input_folder', './inbox'), '_temp_attachments')
        if os.path.exists(temp_dir):
            self.log("Cleaning up temporary attachments...")
            # Move all temp attachments to processed
            for attachment_file in os.listdir(temp_dir):
                src = os.path.join(temp_dir, attachment_file)
                if os.path.isfile(src):
                    self.move_to_processed(src)
            # Remove temp directory
            try:
                os.rmdir(temp_dir)
            except:
                pass

        # Move .msg files to processed
        for msg_path in msg_files_to_move:
            self.move_to_processed(msg_path)

        # Update tracker
        self.log("Updating tracker...")
        tracker_file = self.update_tracker(results)

        # Print summary
        self.log("=" * 60)
        self.log("SUMMARY")
        self.log("=" * 60)

        passed = sum(1 for r in results if r['review_status'] == 'Passed')
        needs_review = sum(1 for r in results if r['review_status'] == 'Needs Review')
        errors = sum(1 for r in results if r['review_status'] == 'Error')

        model = self.config['API'].get('model', 'claude-sonnet-4-5-20250929')
        model_name = self.get_model_display_name(model)

        self.log(f"Documents processed: {len(results)}")
        self.log(f"  ✓ Passed: {passed}")
        self.log(f"  ⚠ Needs Review: {needs_review}")
        self.log(f"  ✗ Errors: {errors}")
        self.log("")
        self.log(f"Total tokens used: {total_tokens:,}")
        self.log(f"Total cost: ${total_cost:.4f}")
        self.log("")
        self.log(f"This analysis used {total_tokens:,} tokens and cost ${total_cost:.4f} using {model_name}")
        self.log("")
        self.log(f"Results saved to: {tracker_file}")
        self.log(f"Log saved to: {log_file}")


def main():
    """Entry point for the script."""
    try:
        pipeline = DocReviewPipeline()
        pipeline.run()
    except KeyboardInterrupt:
        print("\n\nProcess interrupted by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\n\nFATAL ERROR: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
