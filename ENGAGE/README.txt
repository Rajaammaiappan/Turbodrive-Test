Drop your logo image files in this folder, using these exact names,
and NGAGE will automatically embed them in the header (no internet
needed at runtime, no broken images if the corporate machine is
offline):

    rr_logo.png      (or rr_logo.svg)   <- Rolls-Royce logo
    alten_logo.png   (or alten_logo.svg) <- ALTEN logo

If a file is missing, the header falls back to the plain "RR" / "ALTEN"
text badges automatically, so the app will never break because of a
missing logo file.

How to get the files from the links you shared:
  1. Open each link in a browser:
     - ALTEN: https://www.alten.com/wp-content/uploads/2019/01/favicon-alten.png
     - Rolls-Royce: https://www.rolls-royce.com/~/media/Images/R/Rolls-Royce/logo/rr-logo-svg.svg
  2. Right-click -> "Save image as..."
  3. Save them into this "logos" folder with the exact names above
     (rename them if needed).
  4. Restart NGAGE (or just refresh the browser tab) - the real logos
     will appear in the header automatically.

Note: these were fetched over the public internet to your browser, not
by the app itself. The app never makes outbound calls to alten.com or
rolls-royce.com - the images are read from disk and embedded as
base64 once, so NGAGE keeps working even with no internet access.
