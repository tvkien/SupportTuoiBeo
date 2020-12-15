﻿// <auto-generated />
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Migrations;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using SupportTuoiBeo.Data.Contexts;

namespace SupportTuoiBeo.Data.Migrations
{
    [DbContext(typeof(ApplicationDbContext))]
    [Migration("20201215072404_UpdateConstrain")]
    partial class UpdateConstrain
    {
        protected override void BuildTargetModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .UseIdentityColumns()
                .HasAnnotation("Relational:MaxIdentifierLength", 128)
                .HasAnnotation("ProductVersion", "5.0.1");

            modelBuilder.Entity("SupportTuoiBeo.Data.Entities.UserDetails", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int")
                        .UseIdentityColumn();

                    b.Property<string>("MaKH")
                        .HasColumnType("nvarchar(max)");

                    b.Property<int>("Ngay")
                        .HasColumnType("int");

                    b.Property<int>("Thang")
                        .HasColumnType("int");

                    b.Property<long>("TienThanhToan")
                        .HasColumnType("bigint");

                    b.Property<string>("Tinh")
                        .HasColumnType("nvarchar(max)");

                    b.HasKey("Id");

                    b.ToTable("UserDetails");
                });
#pragma warning restore 612, 618
        }
    }
}
