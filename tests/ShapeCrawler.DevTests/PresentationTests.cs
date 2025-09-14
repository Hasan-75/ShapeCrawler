using System.Globalization;
using System.Reflection;
using DocumentFormat.OpenXml.Presentation;
using Fixture;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using ShapeCrawler.Presentations;

namespace ShapeCrawler.DevTests;

public class PresentationTests : SCTest
{
    private readonly Fixtures fixtures = new();

    [Test]
    public void SlideWidth_Getter_returns_presentation_Slides_Width()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));

        // Act & Assert
        pres.SlideWidth.Should().Be(720);
    }

    [Test]
    public void SlideWidth_Setter_sets_presentation_Slides_Width()
    {
        // Arrange
        var pptx = TestAsset("009_table.pptx");
        var pres = new Presentation(pptx);

        // Act
        pres.SlideWidth = 1000;

        // Assert
        pres.SlideWidth.Should().Be(1000);
    }

    [Test]
    public void SlideHeight_Getter_returns_presentation_Slides_Height()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));

        // Act & Assert
        pres.SlideHeight.Should().Be(405);
    }

    [Test]
    public void SlideHeight_Setter_sets_presentation_Slides_Height()
    {
        // Arrange
        var pptx = TestAsset("009_table.pptx");
        var pres = new Presentation(pptx);

        // Act
        pres.SlideHeight = 700;

        // Assert
        pres.SlideHeight.Should().Be(700);
    }

    [Test]
    public void Slides_Count_returns_One_When_presentation_contains_one_slide()
    {
        // Arrange
        var pres17 = new Presentation(TestAsset("017.pptx"));
        var pres16 = new Presentation(TestAsset("016.pptx"));
        var pres75 = new Presentation(TestAsset("075.pptx"));

        // Act & Assert
        pres17.Slides.Count.Should().Be(1);
        pres16.Slides.Count.Should().Be(1);
        pres75.Slides.Count.Should().Be(1);
    }

    [Test]
    public void Slides_Count()
    {
        // Arrange
        var pres = new Presentation(TestAsset("007_2 slides.pptx"));
        var removingSlide = pres.Slides[0];
        var slides = pres.Slides;

        // Act
        removingSlide.Remove();

        // Assert
        slides.Count.Should().Be(1);
    }

    [Test]
    public void Slides_Add_adds_specified_slide_at_the_end_of_slide_collection()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("001.pptx")).Slides[0];
        var destPre = new Presentation(TestAsset("002.pptx"));
        var originSlidesCount = destPre.Slides.Count;
        var expectedSlidesCount = ++originSlidesCount;
        MemoryStream savedPre = new();

        // Act
        destPre.Slides.Add(sourceSlide);

        // Assert
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");

        destPre.Save(savedPre);
        destPre = new Presentation(savedPre);
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");
    }

    [Test]
    public void Slides_Add_should_copy_only_layout_of_copying_slide()
    {
        // Arrange
        var sourcePres = new Presentation(TestAsset("pictures-case004.pptx"));
        var copyingSlide = sourcePres.Slides[0];
        var destPres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var expectedCount = destPres.Slides.Count + 1;
        MemoryStream savedPre = new();

        // Act
        destPres.Slides.Add(copyingSlide);

        // Assert
        destPres.Slides.Count.Should().Be(expectedCount);

        destPres.Save(savedPre);
        destPres = new Presentation(savedPre);
        destPres.Slides.Count.Should().Be(expectedCount);
        destPres.Slides[1].SlideLayout.SlideMaster.SlideLayouts.Count().Should().Be(1);
        destPres.Validate();
    }

    [Test]
    public void Slides_Add_should_copy_notes()
    {
        // Arrange
        var sourcePres = new Presentation(TestAsset("008.pptx"));
        var copyingSlide = sourcePres.Slides[0];
        var destPres = new Presentation(TestAsset("autoshape-case017_slide-number.pptx"));

        // Act
        destPres.Slides.Add(copyingSlide);

        // Assert
        destPres.Slides.Last().Notes!.Text.Should().Be("Test note");
        destPres.Validate();
    }

    [Test]
    public void Slides_Add_adds_a_new_slide()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMaster(1).SlideLayouts.First(l => l.Name == "Blank");

        // Act
        pres.Slides.Add(layout.Number);

        // Assert
        pres.Slides.Count.Should().Be(1);
        pres.Validate();
    }

    [Test]
    public void Slides_Add_adds_a_new_slide_using_blank_layout()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMaster(1).SlideLayout("Blank");
        var stream = new MemoryStream();

        // Act
        pres.Slides.Add(layout.Number);

        // Assert
        pres.Save(stream);
        new Presentation(stream).Slide(1).Shapes.Should().BeEmpty();
    }

    [Test]
    public void Slides_Add_adds_a_new_slide_using_layout_with_one_shape()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMaster(1).SlideLayout(6);
        var stream = new MemoryStream();

        // Act
        pres.Slides.Add(layout.Number);

        // Assert
        pres.Save(stream);
        new Presentation(stream).Slide(1).Shapes.Count.Should().Be(1);
    }

    [Test]
    public void Slides_Add_should_not_break_hyperlink()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case018_rotation.pptx"));
        var inserting = pres.Slide(1);

        // Act
        pres.Slides.Add(inserting, 2);

        // Assert
        pres.Validate();
    }

    [Test]
    public void SlideMastersCount_ReturnsNumberOfMasterSlidesInThePresentation()
    {
        // Arrange
        var pres1 = new Presentation(TestAsset("001.pptx"));
        var pres2 = new Presentation(TestAsset("002.pptx"));

        // Act
        var slideMastersCountCase1 = pres1.SlideMasters.Count();
        var slideMastersCountCase2 = pres2.SlideMasters.Count();

        // Assert
        slideMastersCountCase1.Should().Be(1);
        slideMastersCountCase2.Should().Be(2);
    }

    [Test]
    public void SlideMaster_Shapes_Count_returns_number_of_master_shapes()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);

        // Act
        var masterShapesCount = pres.SlideMasters[0].Shapes.Count;

        // Assert
        masterShapesCount.Should().Be(7);
    }

    [Test]
    public void Sections_Remove_removes_specified_section()
    {
        // Arrange
        var pptxStream = TestAsset("autoshape-case017_slide-number.pptx");
        var pres = new Presentation(pptxStream);
        var removingSection = pres.Sections[0];

        // Act
        pres.Sections.Remove(removingSection);

        // Assert
        pres.Sections.Count.Should().Be(0);
    }

    [Test]
    public void Sections_Remove_should_remove_section_after_Removing_Slide_from_section()
    {
        // Arrange
        var pptxStream = TestAsset("autoshape-case017_slide-number.pptx");
        var pres = new Presentation(pptxStream);
        var removingSection = pres.Sections[0];

        // Act
        pres.Slides[0].Remove();
        pres.Sections.Remove(removingSection);

        // Assert
        pres.Sections.Count.Should().Be(0);
    }

    [Test]
    public void Sections_Section_Slides_Count_returns_Zero_When_section_is_Empty()
    {
        // Arrange
        var pptxStream = TestAsset("008.pptx");
        var pres = new Presentation(pptxStream);
        var section = pres.Sections.GetByName("Section 2");

        // Act
        var slidesCount = section.Slides.Count;

        // Assert
        slidesCount.Should().Be(0);
    }

    [Test]
    public void Sections_Section_Slides_Count_returns_number_of_slides_in_section()
    {
        var pptxStream = TestAsset("autoshape-case017_slide-number.pptx");
        var pres = new Presentation(pptxStream);
        var section = pres.Sections.GetByName("Section 1");

        // Act
        var slidesCount = section.Slides.Count;

        // Assert
        slidesCount.Should().Be(1);
    }

    [Test]
    public void Save_saves_presentation_opened_from_Stream_when_it_was_Saved()
    {
        // Arrange
        var presStream = TestAsset("autoshape-case003.pptx");
        var pres = new Presentation(presStream);
        var textBox = pres.Slides[0].Shapes.Shape<IShape>("AutoShape 2").TextBox!;
        textBox.SetText("Test");

        // Act
        pres.Save();

        // Assert
        pres = new Presentation(presStream);
        textBox = pres.Slides[0].Shapes.Shape<IShape>("AutoShape 2").TextBox!;
        textBox.Text.Should().Be("Test");
    }

    [Test]
    public void Save_should_not_throw_exception()
    {
        var presBytes = TestAsset("001.pptx").ToArray();
        var nonExpandableStream = new MemoryStream(presBytes);
        var pres = new Presentation(nonExpandableStream);
        var outputStream = new MemoryStream();

        // Act
        var saving = () => pres.Save(outputStream);

        // Assert
        saving.Should().NotThrow<Exception>();
    }

    [Test]
    public void Save_sets_the_date_of_the_last_modification()
    {
        // Arrange
        var expectedCreated = DateTime.Parse("2024-01-01T12:34:56Z", CultureInfo.InvariantCulture);
        SCSettings.TimeProvider = new FakeTimeProvider(expectedCreated);
        var pres = new Presentation();
        var expectedModified = DateTime.Parse("2024-02-02T15:30:45Z", CultureInfo.InvariantCulture);
        SCSettings.TimeProvider = new FakeTimeProvider(expectedModified);
        var stream = new MemoryStream();

        // Act
        pres.Save(stream);

        // Assert
        stream.Position = 0;
        var updatedPres = new Presentation(stream);
        updatedPres.Properties.Modified.Should().Be(expectedModified);
    }

    [Test]
    public void Footer_AddSlideNumber_adds_slide_number()
    {
        // Arrange
        var pres = new Presentation(pres => { pres.Slide(); });

        // Act
        pres.Footer.AddSlideNumber();

        // Assert
        pres.Footer.SlideNumberAdded().Should().BeTrue();
    }

    [Test, Ignore("In Progress #540")]
    public void Footer_RemoveSlideNumber_removes_slide_number()
    {
        // Arrange
        var pres = new Presentation();
        pres.Footer.AddSlideNumber();

        // Act
        pres.Footer.RemoveSlideNumber();

        // Assert
        pres.Footer.SlideNumberAdded().Should().BeFalse();
    }

    [Test]
    public void Footer_SlideNumberAdded_returns_false_When_slide_number_is_not_added()
    {
        // Arrange
        var pres = new Presentation();

        // Act-Assert
        pres.Footer.SlideNumberAdded().Should().BeFalse();
    }

    [Test]
    public void Footer_AddText_adds_text_footers_in_all_slides()
    {
        // Arrange
        var pres = new Presentation(p =>
        { 
            p.Slide();
            p.Slide();
        });

        var text = "To infinity and beyond!";

        // Act
        pres.Footer.AddText(text);

        // Assert
        pres.Slides.Should().AllSatisfy(slide =>
        {
            slide.Shapes
                .Should()
                .Contain(shape =>
                    shape.PlaceholderType == PlaceholderType.Footer
                    && shape.TextBox.Text == text);
        });

    }

    [Test]
    public void Footer_RemoveText_removes_text_footers_from_all_slides()
    {
        // Arrange
        var pres = new Presentation(p =>
        { 
            p.Slide();
            p.Slide();
        });

        var text = "To infinity and beyond!";

        // Act
        pres.Footer.AddText(text);
        pres.Footer.RemoveText();

        // Assert
        pres.Slides.Should().AllSatisfy(slide =>
        {
            slide.Shapes
                .Should()
                .NotContain(shape => shape.PlaceholderType == PlaceholderType.Footer);
        });
    }

    [Test]
    public void Footer_AddTextOnSlide_adds_text_footer_on_specific_slide()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide();
            p.Slide();
        });

        var text = "To infinity and beyond";

        // Act
        pres.Footer.AddTextOnSlide(2, text);

        var addInOutOfRangeSlide = () => pres.Footer.AddTextOnSlide(3, "Ow snap!");

        // Assert
        pres.Slides[1].Shapes.Should().Contain(s => s.PlaceholderType == PlaceholderType.Footer && s.TextBox.Text == text);
        pres.Slides[0].Shapes.Should().NotContain(s => s.PlaceholderType == PlaceholderType.Footer);

        addInOutOfRangeSlide.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Test]
    public void Footer_RemoveTextOnSlide_removes_text_footer_from_specific_slide()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide();
            p.Slide();
        });
        var text = "To infinity and beyond!";
        pres.Footer.AddText(text);

        // Act
        pres.Footer.RemoveTextOnSlide(1);
        var removeFromOutOfRangeSlide = () => pres.Footer.RemoveTextOnSlide(0);

        // Assert
        pres.Slides[0].Shapes.Should().NotContain(s => s.PlaceholderType == PlaceholderType.Footer);
        pres.Slides[1].Shapes.Should().Contain(s => s.PlaceholderType == PlaceholderType.Footer && s.TextBox.Text == text);

        removeFromOutOfRangeSlide.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Test]
    public void Slides_Add_adds_slide()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("001.pptx")).Slides[0];
        var pptx = TestAsset("002.pptx");
        var destPre = new Presentation(pptx);
        var originSlidesCount = destPre.Slides.Count;
        var expectedSlidesCount = ++originSlidesCount;
        MemoryStream savedPre = new();

        // Act
        destPre.Slides.Add(sourceSlide);

        // Assert
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");

        destPre.Save(savedPre);
        destPre = new Presentation(savedPre);
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");
    }
    
    [Test]
    public void Slides_Add_copies_slide_with_chart()
    {
        // Arrange
        var file = fixtures.AssemblyStream("084 charts.pptx");
        var pres = new Presentation(file);
        var slide = pres.Slides.Last();

        // Act
        pres.Slides.Add(slide, 1);

        // Assert
        pres.Validate();
    }


    [Test]
    public void Slides_Add_adds_slide_at_position()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide();
            p.Slide(s =>
            {
                s.Picture(
                    "Picture 1",
                    fixtures.Int(),
                    fixtures.Int(),
                    fixtures.Int(),
                    fixtures.Int(),
                    fixtures.Image());
            });
        });
        var copyingSlide = pres.Slide(2);

        // Act
        pres.Slides.Add(copyingSlide, 1);

        // Assert
        pres.Slide(1).Shapes.Shape("Picture 1").Picture.Image!.AsByteArray().Should().NotBeEmpty();
    }

    [Test]
    [TestCase("007_2 slides.pptx", 1)]
    [TestCase("006_1 slides.pptx", 0)]
    public void Slides_Remove_removes_slide(string file, int expectedSlidesCount)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var removingSlide = pres.Slides[0];
        var mStream = new MemoryStream();

        // Act
        removingSlide.Remove();

        // Assert
        pres.Slides.Should().HaveCount(expectedSlidesCount);

        pres.Save(mStream);
        pres = new Presentation(mStream);
        pres.Slides.Should().HaveCount(expectedSlidesCount);
    }

    [Test]
    public void Slides_Add_adds_slide_at_the_specified_position()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("001.pptx")).Slide(1);
        var sourceSlideId = Guid.NewGuid().ToString();
        sourceSlide.CustomData = sourceSlideId;
        var destPres = new Presentation(TestAsset("002.pptx"));

        // Act
        destPres.Slides.Add(sourceSlide, 2);

        // Assert
        destPres.Slide(2).CustomData.Should().Be(sourceSlideId);
    }

    [Test]
    public void Slides_Add_adds_a_new_slide_at_the_specified_position_using_specified_layout()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide(s =>
            {
                s.TextBox(fixtures.String(), fixtures.Int(), fixtures.Int(), fixtures.Int(), fixtures.Int(),
                    fixtures.String());
            });
        });
        var layoutNumber = pres.SlideMasters.Select(sm => sm.SlideLayout("Blank")).First().Number;

        // Act
        pres.Slides.Add(layoutNumber, 1);

        // Assert
        pres.Slide(2).Shapes.Count.Should().Be(1);
    }


    [Test]
    public void FileProperties_Title_Setter_sets_title()
    {
        // Arrange
        var pres = new Presentation();
        var expectedCreated = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Local);

        // Act
        pres.Properties.Title = "Properties_setter_sets_values";
        pres.Properties.Created = expectedCreated;

        // Assert
        pres.Properties.Title.Should().Be("Properties_setter_sets_values");
        pres.Properties.Created.Should().Be(expectedCreated);
    }

    [Test]
    public void FileProperties_getters_return_valid_values_after_saving_presentation()
    {
        // Arrange
        var pres = new Presentation();
        var expectedCreated = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Local);
        var stream = new MemoryStream();

        // Act
        pres.Properties.Title = "Properties_setter_survives_round_trip";
        pres.Properties.Created = expectedCreated;
        pres.Properties.RevisionNumber = 100;
        pres.Save(stream);

        // Assert
        stream.Position = 0;
        var updatePres = new Presentation(stream);
        updatePres.Properties.Title.Should().Be("Properties_setter_survives_round_trip");
        updatePres.Properties.Created.Should().Be(expectedCreated);
        pres.Properties.RevisionNumber.Should().Be(100);
    }

    [Test]
    public void FileProperties_Modified_Getter_returns_date_of_the_last_modification()
    {
        var pres = new Presentation(TestAsset("059_crop-images.pptx"));
        var expectedModified = DateTime.Parse("2024-12-16T17:11:58Z", CultureInfo.InvariantCulture);

        // Act-Assert
        pres.Properties.Modified.Should().Be(expectedModified);
        pres.Properties.Title.Should().Be("");
        pres.Properties.RevisionNumber.Should().Be(7);
        pres.Properties.Comments.Should().BeNull();
    }

    [Test]
    public void Non_parameter_constructor_sets_the_date_of_the_last_modification()
    {
        // Arrange
        var expectedModified = DateTime.Parse("2024-01-01T12:34:56Z", CultureInfo.InvariantCulture);
        SCSettings.TimeProvider = new FakeTimeProvider(expectedModified);

        // Act
        var pres = new Presentation();

        // Assert
        pres.Properties.Modified.Should().Be(expectedModified);
    }

    [Test]
    public void Constructor_does_not_throw_exception_When_the_specified_file_is_a_google_slide_export()
    {
        // Act
        var openingGoogleSlides = () => new Presentation(TestAsset("074 google slides.pptx"));

        // Assert
        openingGoogleSlides.Should().NotThrow();
    }

    [Test]
    public void Create_creates_new_presentation_with_slide()
    {
        // Arrange
        var imageStream = TestAsset("reference image.png");

        // Act
        var pres = new Presentation(pres =>
        {
            pres.Slide(slide =>
            {
                slide.Picture(
                    name: "Picture",
                    x: 100,
                    y: 100,
                    width: 200,
                    height: 50,
                    image: imageStream);
            });
        });

        // Assert
        pres.Slide(1).Shape("Picture").Should().NotBeNull();
    }

    [Test]
    public void Constructor_creates_presentation()
    {
        // Act
        var pres = new Presentation();

        // Assert
        pres.Should().NotBeNull();
        pres.Validate();
    }

    [Test]
    public void Constructor_creates_empty_presentation()
    {
        // Act-Assert
        new Presentation().Slides.Should().BeEmpty();
    }

    [Test]
    public void AsMarkdown_returns_markdown_string()
    {
        // Arrange
        var pres = new Presentation(TestAsset("076 bitcoin.pptx"));
        var expectedMarkdown = StringOf("076 bitcoin.md");

        // Act
        var actualMarkdown = pres.AsMarkdown();

        // Assert
        actualMarkdown.Should().BeEquivalentTo(expectedMarkdown);
    }

    [Test]
    public void Save_does_not_throw_exception_When_stream_is_a_File_stream()
    {
        // Arrange
        var pres = new Presentation();
        var file = Path.GetTempFileName();
        using var stream = File.OpenWrite(file);

        // Act
        var saving = () => pres.Save(stream);

        // Assert
        saving.Should().NotThrow();

        // Cleanup
        stream.Close();
        File.Delete(file);
    }

    [Test]
    public void Slides_RemoveThenAdd_EnsuresUniqueSlideIds()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("001.pptx")).Slide(1);
        var destPres = new Presentation(TestAsset("001.pptx"));

        // Act
        destPres.Slide(2).Remove();
        destPres.Slides.Add(sourceSlide, 1);

        // Assert
        var slideIdRelationshipIdList =
            destPres.GetSDKPresentationDocument().PresentationPart!.Presentation.SlideIdList!.OfType<SlideId>()
                .Select(s => s.RelationshipId);
        slideIdRelationshipIdList.Should().OnlyHaveUniqueItems();
    }

    [Test]
    public void Slide_throws_exception()
    {
        // Arrange
        var pres = new Presentation();

        // Act
        var accessUnavailableSlide = () => pres.Slide(1);

        // Assert
        accessUnavailableSlide.Should().Throw<Exception>();
    }
}