﻿using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace Hammer.MDI.Control.WindowControls
{
    internal class HollowRectangleAdorner : Adorner
    {
        public HollowRectangleAdorner(UIElement adornedElement)
            : base(adornedElement)
        {
        }

        protected override Size MeasureOverride(Size constraint)
        {
            var result = base.MeasureOverride(constraint);

            InvalidateVisual();
            return result;
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            if ((AdornedElement as MdiWindow)?.Container == null)
            {
                return;
            }

            var renderBrush = new SolidColorBrush(Colors.Transparent)
            {
                Opacity = 0.1
            };
            var renderPen = new Pen(new SolidColorBrush(Colors.Transparent), 0);

            var x = Math.Max(0, Canvas.GetLeft(AdornedElement));
            var y = Math.Max(0, Canvas.GetTop(AdornedElement));
            var w = AdornedElement.DesiredSize.Width;
            var h = AdornedElement.DesiredSize.Height;
            double ww = 0;
            double hh = 0;
            if (AdornedElement != null && AdornedElement is MdiWindow adornerWindow && adornerWindow.Container != null)
            {
                ww = adornerWindow.Container.ActualWidth;
                hh = adornerWindow.Container.ActualHeight;
            }
            var xPlusw = Math.Min(ww, x + w);
            var yPlush = Math.Min(hh, y + h);

            var pointA = new Point(0, 0);
            var pointB = new Point(xPlusw, 0);
            var pointD = new Point(0, y);
            var pointF = new Point(xPlusw, y);
            var pointG = new Point(0, yPlush);
            var pointH = new Point(x, yPlush);
            var pointK = new Point(xPlusw, hh);
            var pointL = new Point(ww, hh);

            var rectangle1 = new Rect(pointA, pointF);
            var rectangle2 = new Rect(pointB, pointL);
            var rectangle3 = new Rect(pointG, pointK);
            var rectangle4 = new Rect(pointD, pointH);

            drawingContext.PushTransform(new TranslateTransform(-Canvas.GetLeft(AdornedElement), -Canvas.GetTop(AdornedElement)));
            drawingContext.DrawRectangle(renderBrush, renderPen, rectangle1);
            drawingContext.DrawRectangle(renderBrush, renderPen, rectangle2);
            drawingContext.DrawRectangle(renderBrush, renderPen, rectangle3);
            drawingContext.DrawRectangle(renderBrush, renderPen, rectangle4);
        }
    }
}