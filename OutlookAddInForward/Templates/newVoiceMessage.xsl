<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl">
  <xsl:output method="html"
              doctype-system="about:legacy-compat"
              encoding="UTF-8"
              indent="yes"/>

  <xsl:template match="properties">
    <html lang="en">
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <style type="text/css">
          /*
          * What follows is the result of much research on cross-browser styling.
          * Credit left inline and big thanks to Nicolas Gallagher, Jonathan Neal,
          * Kroc Camen, and the H5BP dev community and team.
          */

          /* ==========================================================================
          Base styles: opinionated defaults
          ========================================================================== */

          html {
          color: #222;
          font-size: 1em;
          line-height: 1.4;
          }

          /*
          * Remove text-shadow in selection highlight:
          * https://twitter.com/miketaylr/status/12228805301
          *
          * These selection rule sets have to be separate.
          * Customize the background color to match your design.
          */

          ::-moz-selection {
          background: #b3d4fc;
          text-shadow: none;
          }

          ::selection {
          background: #b3d4fc;
          text-shadow: none;
          }

          /*
          * A better looking default horizontal rule
          */

          hr {
          display: block;
          height: 1px;
          border: 0;
          border-top: 1px solid #ccc;
          margin: 1em 0;
          padding: 0;
          }

          /*
          * Remove the gap between audio, canvas, iframes,
          * images, videos and the bottom of their containers:
          * https://github.com/h5bp/html5-boilerplate/issues/440
          */

          audio,
          canvas,
          iframe,
          img,
          svg,
          video {
          vertical-align: middle;
          }

          /*
          * Remove default fieldset styles.
          */

          fieldset {
          border: 0;
          margin: 0;
          padding: 0;
          }

          /*
          * Allow only vertical resizing of textareas.
          */

          textarea {
          resize: vertical;
          }

          /* ==========================================================================
          Browser Upgrade Prompt
          ========================================================================== */

          .browserupgrade {
          margin: 0.2em 0;
          background: #ccc;
          color: #000;
          padding: 0.2em 0;
          }

          /* ==========================================================================
          Author's custom styles
          ========================================================================== */

          .alert-section {
          color: #5f5f5f;
          font-family: "Arial","sans-serif";
          font-size: 10pt;
          width: 710px;
          margin: 20px;
          }

          .alert-head .alert-title {
          font-size: 18pt;
          color: #231c1c;
          font-family: "Arial","sans-serif";
          }

          .alert-body {
          padding-right: 55px;
          }

          .alert-body p {
          text-align: justify;
          }

          div.divider {
          margin-top: 20px;
          padding-top: 8px;
          text-align: right;
          font-size: 0.8em;
          color: #939292;
          border-top: 2px solid #ebb700;
          }


          /* ==========================================================================
          Helper classes
          ========================================================================== */
          /*
          * Hide visually and from screen readers:
          * http://juicystudio.com/article/screen-readers-display-none.php
          */
          .hidden {
          display: none !important;
          visibility: hidden;
          }

          /*
          * Hide only visually, but have it available for screen readers:
          * http://snook.ca/archives/html_and_css/hiding-content-for-accessibility
          */

          .visuallyhidden {
          border: 0;
          clip: rect(0 0 0 0);
          height: 1px;
          margin: -1px;
          overflow: hidden;
          padding: 0;
          position: absolute;
          width: 1px;
          }

          /*
          * Extends the .visuallyhidden class to allow the element
          * to be focusable when navigated to via the keyboard:
          * https://www.drupal.org/node/897638
          */

          .visuallyhidden.focusable:active,
          .visuallyhidden.focusable:focus {
          clip: auto;
          height: auto;
          margin: 0;
          overflow: visible;
          position: static;
          width: auto;
          }

          /*
          * Hide visually and from screen readers, but maintain layout
          */

          .invisible {
          visibility: hidden;
          }

          /*
          * Clearfix: contain floats
          *
          * For modern browsers
          * 1. The space content is one way to avoid an Opera bug when the
          *    `contenteditable` attribute is included anywhere else in the document.
          *    Otherwise it causes space to appear at the top and bottom of elements
          *    that receive the `clearfix` class.
          * 2. The use of `table` rather than `block` is only necessary if using
          *    `:before` to contain the top-margins of child elements.
          */

          .clearfix:before,
          .clearfix:after {
          content: " "; /* 1 */
          display: table; /* 2 */
          }

          .clearfix:after {
          clear: both;
          }
        </style>
      </head>
      <body>
        <table border="0" cellspacing="0" cellpadding="0" width="700" style='width:700px;'>
          <tr>
            <td>
              <div class="alert-section">
                <div class="alert-head">
                  <p class="alert-title">Notificación de Mensaje de Voz</p>
                </div>
                <div class="alert-body">
                  <p>
                    Estimado <strong>
                      <xsl:value-of select="usuarioAnexo"/>
                    </strong>,
                  </p>
                  <div>
                    <p>
                      Le informamos que usted ha recibido un mensaje de voz del número
                      <xsl:choose>
                        <xsl:when test="llamadaInterna = 1">
                          de anexo <strong>
                            <xsl:value-of select="numeroTelefono"/>
                          </strong>
                        </xsl:when>
                        <xsl:otherwise>
                          externo <strong>
                            <xsl:value-of select="numeroTelefono"/>
                          </strong>
                        </xsl:otherwise>
                      </xsl:choose>. A continuación indicamos algunos detalles del mensaje de voz:
                    </p>
                    <ul style="margin-bottom:30px;">
                      <li>
                        Anexo de recepción:  <xsl:value-of select="anexo"/>
                      </li>
                      <li>
                        Número de teléfono: <xsl:value-of select="numeroTelefono"/>
                        <xsl:choose>
                          <xsl:when test="llamadaInterna = 1">
                            (<strong>anexo interno</strong>)
                          </xsl:when>
                          <xsl:otherwise>
                            (<strong>número externo</strong>)
                          </xsl:otherwise>
                        </xsl:choose>
                      </li>
                      <li>
                        Fecha del mensaje: <xsl:value-of select="fechaLlamada"/>
                      </li>
                      <li>
                        Hora del mensaje: <xsl:value-of select="horaLlamada"/>
                      </li>
                    </ul>
                    <p>
                      Para escuchar los mensajes de voz en su anexo telefónico de la oficina mientras usted se encuentra fuera de ésta, llame al número de la central 6188500 marcar 777 luego ingresar su número de anexo, su contraseña y escoja la opción 11 para escuchar sus mensajes almacenados.
                    </p>
                    <p>
                      Por favor encuentre adjunto el mensaje de voz recibido.
                    </p>
                    <p>
                      Atentamente.
                    </p>
                  </div>

                  <p style="margin-top:30px;">
                    Departamento de Sistemas
                  </p>
                  <div class="divider">
                    Este es un mensaje del Sistema de Notificaciones de la Central Telefónica.
                  </div>
                </div>
              </div>
            </td>
          </tr>
        </table>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>
