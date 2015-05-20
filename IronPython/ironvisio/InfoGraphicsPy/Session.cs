﻿using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;
using IVisio=Microsoft.Office.Interop.Visio;
using IG=InfoGraphicsPy;
using System.Linq;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace InfoGraphicsPy
{
    public class Session
    {
        private IVisio.Application app;
        private IVisio.Document doc;
        private IVisio.Document stencil;

        private IVisio.Master _masterRectangle;
        
        public Session()
        {
            this.app = new IVisio.Application();
            this.NewDocument();
        }

        public void NewDocument()
        {
            var docs = this.Application.Documents;
            this.doc = docs.Add("");
            this.stencil = docs.OpenStencil("basic_u.vss");
            var masters = stencil.Masters;
            this._masterRectangle = masters["Rectangle"];
        }

        public void NewDocument(double w, double h)
        {
            var docs = this.Application.Documents;
            this.doc = docs.Add("");
        }

        public void NewPage()
        {
            var doc = this.doc;
            doc.Pages.Add();
        }

        public void ResizePageToFit()
        {
            var page = this.Page;
            page.ResizeToFitContents();
        }

        public void ResizePageToFit(double w, double h)
        {
            var page = this.Page;
            page.ResizeToFitContents(new VA.Drawing.Size(w,h));
        }

        public IVisio.Application Application
        {
            get { return this.app; }
        }

        public IVisio.Page Page
        {
            get { return this.Application.ActivePage; }
        }

        public Master MasterRectangle
        {
            get { return _masterRectangle; }
        }
    }
}
