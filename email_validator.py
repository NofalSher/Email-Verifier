"""
Email Validator - A tool for validating email addresses using SMTP verification.

This module provides functionality to validate email addresses by checking their format,
domain MX records, and SMTP server responses.
"""

import smtplib
import dns.resolver
import logging
import time
import random
import pandas as pd
from typing import List, Tuple, Optional
from email_validator import validate_email, EmailNotValidError


class EmailValidator:
    """A class for validating email addresses using SMTP verification."""
    
    def __init__(self, log_file: str = 'email_verification.log', debug_level: int = 0):
        """
        Initialize the EmailValidator.
        
        Args:
            log_file (str): Path to the log file
            debug_level (int): SMTP debug level (0 = no debug, 1 = verbose)
        """
        self.log_file = log_file
        self.debug_level = debug_level
        self.sender_email = 'verification@example.com'
        self._setup_logging()
    
    def _setup_logging(self) -> None:
        """Configure logging for the application."""
        logging.basicConfig(
            filename=self.log_file,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            filemode='a'
        )
    
    def get_mx_record(self, domain: str) -> Optional[str]:
        """
        Get the MX record of the domain.
        
        Args:
            domain (str): The domain to lookup
            
        Returns:
            Optional[str]: The MX record or None if not found
        """
        try:
            mx_records = dns.resolver.resolve(domain, 'MX')
            mx_record = str(mx_records[0].exchange)
            logging.info(f"MX record found for {domain}: {mx_record}")
            return mx_record
        except Exception as e:
            logging.error(f"Error retrieving MX records for {domain}: {e}")
            return None
    
    def verify_email(self, email: str) -> Tuple[bool, str]:
        """
        Verify if an email address exists using SMTP.
        
        Args:
            email (str): Email address to verify
            
        Returns:
            Tuple[bool, str]: (is_valid, message)
        """
        try:
            # Validate email format
            try:
                valid = validate_email(email)
                normalized_email = valid.email
            except EmailNotValidError as e:
                return False, f"Invalid email format: {e}"
            
            # Split email into local part and domain
            local_part, domain = normalized_email.split('@')
            
            # Get MX record
            mx_record = self.get_mx_record(domain)
            if not mx_record:
                return False, "Domain MX record not found"
            
            # Connect to email server
            with smtplib.SMTP(mx_record, timeout=10) as server:
                server.set_debuglevel(self.debug_level)
                server.helo()
                server.mail(self.sender_email)
                code, message = server.rcpt(normalized_email)
                
                if code == 250:
                    return True, "Email address is valid"
                else:
                    return False, f"Server response: {code} {message.decode('utf-8')}"
                    
        except smtplib.SMTPConnectError as e:
            return False, f"SMTP connection error: {e}"
        except smtplib.SMTPServerDisconnected as e:
            return False, f"SMTP server disconnected: {e}"
        except Exception as e:
            return False, f"Error during verification: {e}"
    
    def process_email_list(self, email_list: List[str], 
                          min_delay: float = 1.0, 
                          max_delay: float = 5.0) -> List[Tuple[str, bool, str]]:
        """
        Process a list of emails and return validation results.
        
        Args:
            email_list (List[str]): List of email addresses to validate
            min_delay (float): Minimum delay between requests (seconds)
            max_delay (float): Maximum delay between requests (seconds)
            
        Returns:
            List[Tuple[str, bool, str]]: List of (email, is_valid, message) tuples
        """
        results = []
        total_emails = len(email_list)
        
        for i, email in enumerate(email_list, 1):
            if not isinstance(email, str):
                message = "Invalid email type"
                logging.warning(f"Email: {email}, Valid: False, Message: {message}")
                results.append((email, False, message))
                continue
            
            # Verify the email
            is_valid, message = self.verify_email(email)
            results.append((email, is_valid, message))
            
            # Log and print results
            log_message = f"[{i}/{total_emails}] Email: {email}, Valid: {is_valid}, Message: {message}"
            logging.info(log_message)
            print(log_message)
            
            # Rate limiting with randomized delay (except for last email)
            if i < total_emails:
                delay = random.uniform(min_delay, max_delay)
                time.sleep(delay)
        
        return results


class ExcelHandler:
    """A class for handling Excel file operations."""
    
    @staticmethod
    def read_emails_from_excel(file_path: str, email_column_name: str) -> Tuple[List[str], Optional[pd.DataFrame]]:
        """
        Read emails from an Excel file.
        
        Args:
            file_path (str): Path to the Excel file
            email_column_name (str): Name of the column containing emails
            
        Returns:
            Tuple[List[str], Optional[pd.DataFrame]]: (email_list, dataframe)
        """
        try:
            df = pd.read_excel(file_path)
            
            if email_column_name not in df.columns:
                logging.error(f"Column '{email_column_name}' not found in Excel file")
                return [], None
            
            emails = df[email_column_name].dropna().astype(str).tolist()
            logging.info(f"Successfully read {len(emails)} emails from {file_path}")
            return emails, df
            
        except FileNotFoundError:
            logging.error(f"Excel file not found: {file_path}")
            return [], None
        except Exception as e:
            logging.error(f"Error reading Excel file: {e}")
            return [], None
    
    @staticmethod
    def write_results_to_excel(df: pd.DataFrame, output_file_path: str) -> bool:
        """
        Write the validation results to an Excel file.
        
        Args:
            df (pd.DataFrame): DataFrame to write
            output_file_path (str): Path for the output file
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            df.to_excel(output_file_path, index=False)
            logging.info(f"Results written to Excel file: {output_file_path}")
            return True
        except Exception as e:
            logging.error(f"Error writing to Excel file: {e}")
            return False


def main():
    """Main function to run the email validation process."""
    # Configuration
    INPUT_FILE = 'input_emails.xlsx'  # Change this to your input file
    EMAIL_COLUMN = 'Email'  # Change this to your email column name
    OUTPUT_FILE = 'verified_emails.xlsx'
    
    # Initialize validator
    validator = EmailValidator()
    excel_handler = ExcelHandler()
    
    print("Starting email validation process...")
    
    # Read emails from Excel file
    emails, df = excel_handler.read_emails_from_excel(INPUT_FILE, EMAIL_COLUMN)
    
    if not emails or df is None:
        print("No emails to process. Please check your input file and column name.")
        logging.error("No emails to process.")
        return
    
    print(f"Found {len(emails)} emails to validate.")
    
    # Process the email list
    verification_results = validator.process_email_list(emails)
    
    # Add validation results to DataFrame
    validation_status = [result[1] for result in verification_results]
    validation_messages = [result[2] for result in verification_results]
    
    df['Is_Valid'] = validation_status
    df['Validation_Message'] = validation_messages
    
    # Write results to Excel
    if excel_handler.write_results_to_excel(df, OUTPUT_FILE):
        print(f"Validation complete! Results saved to {OUTPUT_FILE}")
    else:
        print("Error saving results to Excel file.")
    
    # Print summary
    valid_count = sum(validation_status)
    total_count = len(validation_status)
    print(f"\nSummary: {valid_count}/{total_count} emails are valid ({valid_count/total_count*100:.1f}%)")


if __name__ == "__main__":
    main()